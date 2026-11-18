"""Cria requisições PR no GPS usando as linhas da planilha de emissão de NP.

O programa localiza a planilha Excel nesta mesma pasta, apresenta uma prévia
dos registros e, após a confirmação explícita do operador, preenche a tela
Requisition PR (21.3). Cada requisição só é registrada como concluída na
coluna N quando o GPS exibe a mensagem "New requisition created".

Esta automação usa coordenadas fixas da tela e foi preparada para a resolução
1920x1080 com a janela do GPS maximizada. Alterações na resolução, escala do
Windows ou layout da tela podem exigir atualizar as coordenadas abaixo.
"""

from __future__ import annotations

import argparse
import os
import subprocess
import time
from pathlib import Path
from typing import Any

import openpyxl
import pyautogui
import pygetwindow
import pyperclip


GPS_TITLE = "GPS - Production Environment"
# A resolução mínima esperada para que as coordenadas de clique coincidam
# com a tela em que o mapeamento foi feito.
DESIGN_SCREEN_SIZE = (1920, 1080)
# Pequena espera para que a tela de requisição termine de carregar.
PAGE_LOAD_SECONDS = 3

# Coordenadas (x, y) dos campos na tela Requisition PR (21.3), obtidas com
# o GPS maximizado. As chaves correspondem aos cabeçalhos do Excel.
FIELD_COORDINATES = {
    "PN/PHANTOM": (414, 158),
    "V/S": (518, 158),
    "COD GTP": (342, 177),
    "COMPRADOR": (387, 177),
    "USER": (346, 196),
    "REQ WEEK": (387, 215),
    "UOM": (481, 215),
    "REQ MSG": (657, 215),
    "ANUAL(NP)": (400, 234),
    "SUFIXO": (661, 253),
    "ORDER QDE(RB)": (400, 272),
    "CO/BL": (779, 272),
    "NOTES": (612, 405),
}
# Controles da tela que não vêm de colunas do Excel: sinalizar nova requisição,
# iniciar o pedido e acionar a gravação.
CREATE_NEW_COORDINATE = (653, 196)
INITIATE_COORDINATE = (329, 443)
UPDATE_BUTTON_COORDINATE = (405, 581)
# Acesso à tela inicial e ao campo Command para abrir o formulário 21.3.
START_PAGE_COORDINATE = (113, 64)
COMMAND_COORDINATE = (84, 935)
# A ordem dos cabeçalhos é usada para validar a planilha e montar cada registro.
REQUIRED_HEADERS = tuple(FIELD_COORDINATES)
# A coluna N (14ª coluna) guarda o retorno de sucesso do GPS.
STATUS_COLUMN = 14
STATUS_HEADER = "STATUS"

# Script PowerShell embutido que consulta a árvore de acessibilidade do GPS.
# Ele procura a frase de sucesso sem depender do sufixo variável (por exemplo,
# T33312p) e retorna apenas a frase fixa para ser gravada no Excel.
SUCCESS_MESSAGE_SCRIPT = r"""
Add-Type -AssemblyName UIAutomationClient
$deadline = [DateTime]::UtcNow.AddSeconds([int]$env:GPS_STATUS_TIMEOUT)
$previousMessage = [string]$env:GPS_PREVIOUS_SUCCESS
do {
    # O processo pode ainda não estar pronto, então tenta localizá-lo em cada
    # ciclo até o prazo configurado.
    $gps = Get-Process -Name GPSProd -ErrorAction SilentlyContinue |
        Where-Object { $_.MainWindowHandle -ne 0 } |
        Select-Object -First 1
    if ($null -ne $gps) {
        $root = [System.Windows.Automation.AutomationElement]::FromHandle($gps.MainWindowHandle)
        $condition = New-Object System.Windows.Automation.PropertyCondition(
            [System.Windows.Automation.AutomationElement]::ProcessIdProperty,
            $gps.Id
        )
        $elements = $root.FindAll(
            [System.Windows.Automation.TreeScope]::Descendants,
            $condition
        )
        foreach ($element in $elements) {
            $name = $element.Current.Name
            # A tela costuma acrescentar um identificador depois da frase.
            # O identificador é deliberadamente ignorado; basta o sucesso.
            if ($name -match '^\s*New requisition created(?:\s+.*)?$') {
                $message = 'New requisition created'
                if ($message -ne $previousMessage) {
                    $message
                    exit 0
                }
            }
        }
    }
    Start-Sleep -Milliseconds 500
} while ([DateTime]::UtcNow -lt $deadline)
exit 2
"""


def find_workbook(folder: Path) -> Path:
    """Localiza a única planilha GPS de NP com extensão .xlsx na pasta."""
    candidates = [
        path
        for path in folder.glob("*.xlsx")
        if path.name.casefold().startswith("gps -")
        and path.stem.casefold().endswith("np")
    ]
    if len(candidates) != 1:
        # Falhar explicitamente evita escolher silenciosamente a planilha errada
        # quando não há arquivo ou quando há mais de um candidato.
        names = ", ".join(path.name for path in candidates) or "nenhuma"
        raise FileNotFoundError(
            f"Esperava exatamente uma planilha GPS de NP na pasta; encontrei: {names}"
        )
    return candidates[0]


def cell_text(value: Any) -> str:
    """Converte valores do Excel para texto adequado aos campos do GPS."""
    if value is None:
        return ""
    if isinstance(value, bool):
        # Os campos de sinalização do GPS usam Y/N em vez de True/False.
        return "Y" if value else "N"
    if isinstance(value, float) and value.is_integer():
        # Evita enviar códigos/quantidades como "43.0" quando o Excel os lê
        # como números inteiros representados internamente por float.
        return str(int(value))
    return str(value).strip()


def read_records(workbook_path: Path) -> list[dict[str, Any]]:
    """Lê e valida os registros, guardando também o número real da linha Excel.

    O modo somente leitura reduz o trabalho feito na etapa de leitura. O arquivo
    é reaberto mais tarde em modo editável apenas para registrar a confirmação
    de cada requisição na coluna N.
    """
    try:
        workbook = openpyxl.load_workbook(
            workbook_path,
            data_only=True,
            read_only=True,
        )
    except PermissionError as error:
        # O Excel ou a sincronização do OneDrive pode bloquear a abertura.
        raise PermissionError(
            f"Sem acesso à planilha {workbook_path.name}. Feche o arquivo no Excel "
            "e verifique se o OneDrive terminou de sincronizá-lo."
        ) from error
    try:
        sheet = workbook.active
        header_values = next(
            sheet.iter_rows(min_row=1, max_row=1, values_only=True),
            None,
        )
        if header_values is None:
            raise ValueError("A planilha está vazia.")
        headers = [cell_text(value).upper() for value in header_values]
        non_empty_headers = [header for header in headers if header]
        if len(non_empty_headers) != len(set(non_empty_headers)):
            raise ValueError("A planilha contém cabeçalhos duplicados.")

        # O mapeamento é feito pelo texto do cabeçalho, não pela posição da
        # coluna, para que a ordem das colunas de entrada possa mudar.
        missing = [header for header in REQUIRED_HEADERS if header not in headers]
        if missing:
            raise ValueError(
                "Faltam cabeçalhos obrigatórios na planilha: " + ", ".join(missing)
            )

        column_indexes = {header: headers.index(header) for header in REQUIRED_HEADERS}
        records: list[dict[str, Any]] = []
        # Começa na segunda linha porque a primeira contém os nomes das colunas.
        for row_number, values in enumerate(
            sheet.iter_rows(min_row=2, values_only=True),
            start=2,
        ):
            record = {
                header: cell_text(values[column_indexes[header]])
                if column_indexes[header] < len(values)
                else ""
                for header in REQUIRED_HEADERS
            }
            if not any(record.values()):
                continue
            if not record["PN/PHANTOM"]:
                # O código do item identifica a requisição; sem ele, a linha
                # não pode ser preenchida com segurança.
                raise ValueError(
                    f"A linha {row_number} tem dados, mas PN/PHANTOM está vazio."
                )
            record["_excel_row"] = row_number
            records.append(record)

        if not records:
            raise ValueError("A planilha não contém linhas para processar.")
        return records
    finally:
        workbook.close()


def print_preview(records: list[dict[str, Any]]) -> None:
    """Exibe no terminal os valores que serão enviados ao GPS."""
    print(f"\nRegistros encontrados: {len(records)}")
    for row_number, record in enumerate(records, start=1):
        print(f"\nLinha de dados {row_number}:")
        for header in REQUIRED_HEADERS:
            print(f"  {header}: {record[header] or '(vazio)'}")


def read_success_message(
    previous_message: str | None = None,
    timeout_seconds: int = 35,
    allow_missing: bool = False,
) -> str | None:
    """Aguarda e lê a frase de sucesso pela automação de acessibilidade.

    `previous_message` serve para não tratar como nova uma confirmação que
    já estava na tela antes da gravação. Com `allow_missing=True`, a ausência
    de uma mensagem anterior é normal e retorna None; após a gravação, a
    ausência da frase até o prazo gera TimeoutError.
    """
    environment = os.environ.copy()
    environment["GPS_STATUS_TIMEOUT"] = str(timeout_seconds)
    environment["GPS_PREVIOUS_SUCCESS"] = previous_message or ""
    try:
        # O PowerShell consulta UI Automation, pois o texto de status do GPS
        # não é disponibilizado como parte do formulário de campos.
        result = subprocess.run(
            [
                "powershell.exe",
                "-NoProfile",
                "-NonInteractive",
                "-Command",
                SUCCESS_MESSAGE_SCRIPT,
            ],
            env=environment,
            capture_output=True,
            text=True,
            timeout=timeout_seconds + 5,
            check=False,
        )
    except subprocess.TimeoutExpired as error:
        raise TimeoutError(
            "Tempo excedido aguardando a confirmação do GPS."
        ) from error
    if result.returncode == 2:
        if allow_missing:
            return None
        raise TimeoutError(
            "O GPS não exibiu 'New requisition created' após a gravação."
        )
    if result.returncode != 0:
        # Outros códigos indicam falha de execução da consulta, não ausência
        # normal do aviso de sucesso.
        details = result.stderr.strip() or result.stdout.strip()
        raise RuntimeError(
            f"Falha ao consultar o aviso de sucesso do GPS: {details}"
        )
    message = " ".join(result.stdout.split())
    # Normaliza para a frase fixa: qualquer código variável que apareça depois
    # da mensagem original não é usado como parte da validação ou do registro.
    if message != "New requisition created":
        raise RuntimeError(
            f"O GPS retornou uma confirmação inesperada: {message!r}"
        )
    return message


def save_success_status(
    workbook_path: Path,
    worksheet_row: int,
    message: str,
) -> None:
    """Grava a confirmação do GPS na coluna N da linha correspondente."""
    try:
        workbook = openpyxl.load_workbook(workbook_path)
    except PermissionError as error:
        # O GPS pode ter concluído a criação mesmo que o Excel esteja bloqueado.
        # O erro deixa explícito que o status ainda precisa ser registrado.
        raise PermissionError(
            f"GPS confirmou a requisição, mas não foi possível atualizar "
            f"{workbook_path.name}. Feche a planilha no Excel e tente registrar "
            "o status novamente."
        ) from error
    try:
        sheet = workbook.active
        # Cria o rótulo apenas se a célula estiver vazia; não substitui outro
        # cabeçalho que já exista na coluna N.
        if not sheet.cell(row=1, column=STATUS_COLUMN).value:
            sheet.cell(row=1, column=STATUS_COLUMN, value=STATUS_HEADER)
        sheet.cell(row=worksheet_row, column=STATUS_COLUMN, value=message)
        try:
            # Salva após cada registro para não perder confirmações anteriores
            # se uma linha posterior falhar.
            workbook.save(workbook_path)
        except PermissionError as error:
            raise PermissionError(
                f"GPS confirmou a requisição, mas não foi possível salvar "
                f"{workbook_path.name}. Feche a planilha no Excel e tente "
                "registrar o status novamente."
            ) from error
    finally:
        workbook.close()


def get_gps_window() -> Any:
    """Localiza, maximiza e ativa a única janela do GPS."""
    windows = pygetwindow.getWindowsWithTitle(GPS_TITLE)
    if len(windows) != 1:
        raise RuntimeError(
            f"Esperava uma janela GPS com título '{GPS_TITLE}', encontrei {len(windows)}."
        )

    window = windows[0]
    width, height = pyautogui.size()
    # A janela pode ser maximizada, mas a tela física ainda precisa ter
    # resolução suficiente para as coordenadas do formulário.
    if width < DESIGN_SCREEN_SIZE[0] or height < DESIGN_SCREEN_SIZE[1]:
        raise RuntimeError(
            "A tela precisa ter pelo menos 1920x1080 para usar as coordenadas "
            f"mapeadas; resolução atual: {width}x{height}."
        )
    window.maximize()
    window.activate()
    time.sleep(1)
    # Confirma que a maximização realmente ocorreu antes de clicar nos campos.
    if window.width < 1900 or window.height < 1000:
        raise RuntimeError(
            "Não foi possível maximizar a janela GPS; "
            f"tamanho atual: {window.width}x{window.height}."
        )
    return window


def paste_into_field(coordinate: tuple[int, int], value: str) -> None:
    """Substitui o conteúdo de um campo localizado por coordenadas.

    Ctrl+A e Backspace removem qualquer valor anterior; a área de transferência
    permite colar texto de forma consistente, inclusive códigos com zeros à
    esquerda. Se o valor de origem estiver vazio, o campo fica limpo.
    """
    pyautogui.click(*coordinate)
    pyautogui.hotkey("ctrl", "a")
    pyautogui.press("backspace")
    if value:
        pyperclip.copy(value)
        pyautogui.hotkey("ctrl", "v")
    time.sleep(0.1)


def create_record(record: dict[str, Any]) -> str:
    """Preenche o formulário, grava a requisição e aguarda o sucesso do GPS.

    Esta função aciona o botão 5=Upd. Só retorna depois de encontrar a mensagem
    de sucesso, para que a planilha não marque como concluída uma gravação não
    confirmada.
    """
    previous_message = read_success_message(
        timeout_seconds=1,
        allow_missing=True,
    )
    # Sinaliza explicitamente a criação e envia os campos mapeados a partir da
    # linha atual do Excel.
    paste_into_field(CREATE_NEW_COORDINATE, "Y")
    for header, coordinate in FIELD_COORDINATES.items():
        paste_into_field(coordinate, record[header])
    paste_into_field(INITIATE_COORDINATE, "Y")

    # Esta ação grava a requisição no GPS; a resposta visual é verificada logo
    # depois antes que a linha seja marcada como concluída.
    pyautogui.click(*UPDATE_BUTTON_COORDINATE)
    message = read_success_message(previous_message=previous_message)
    if message is None:
        raise RuntimeError("Não foi possível confirmar a gravação no GPS.")
    return message


def open_requisition_page() -> None:
    """Retorna à página inicial e abre novamente Requisition PR (21.3)."""
    pyautogui.click(*START_PAGE_COORDINATE)
    time.sleep(1)
    # O campo Command aceita o identificador da tela e Enter a abre.
    paste_into_field(COMMAND_COORDINATE, "21.3")
    pyautogui.press("enter")
    time.sleep(PAGE_LOAD_SECONDS)


def main() -> None:
    """Executa a prévia, pede confirmação e processa as linhas da planilha."""
    parser = argparse.ArgumentParser(
        description="Preenche e grava novas requisições GPS a partir do Excel."
    )
    parser.add_argument(
        "--dry-run",
        action="store_true",
        help="mostra os dados sem interagir com o GPS",
    )
    args = parser.parse_args()

    # O arquivo Excel esperado fica ao lado deste script.
    workbook_path = find_workbook(Path(__file__).resolve().parent)
    records = read_records(workbook_path)
    print(f"Planilha: {workbook_path.name}")
    print_preview(records)

    if args.dry_run:
        # Modo de conferência: não maximiza o GPS nem interage com ele.
        print("\nPrévia concluída; nenhuma alteração foi feita no GPS.")
        return

    # Exige confirmação exata, incluindo a quantidade, antes de criar registros.
    confirmation = input(
        f"\nIsso criará {len(records)} novas requisições. "
        f"Digite CRIAR {len(records)} para continuar: "
    )
    if confirmation != f"CRIAR {len(records)}":
        print("Operação cancelada; nenhuma alteração foi feita no GPS.")
        return

    # Maximiza/ativa o GPS e navega até a tela de entrada antes do primeiro item.
    get_gps_window()
    open_requisition_page()

    # Processa uma linha de cada vez: grava no GPS, espera o sucesso, salva em N
    # e só então passa à próxima linha.
    for index, record in enumerate(records, start=1):
        print(f"Gravando requisição {index}/{len(records)}...")
        success_message = create_record(record)
        save_success_status(
            workbook_path,
            record["_excel_row"],
            success_message,
        )
        print(
            f"GPS confirmou '{success_message}'; "
            f"status salvo na coluna N, linha {record['_excel_row']}."
        )
        if index < len(records):
            # Cada requisição nova começa em um formulário limpo, reaberto pelo
            # comando da tela 21.3.
            open_requisition_page()

    print("\nProcessamento concluído.")


if __name__ == "__main__":
    # Impede que a automação comece quando o arquivo for importado como módulo
    # durante testes ou reutilização de alguma função.
    main()