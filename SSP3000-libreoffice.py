import winreg
import platform
from odf.opendocument import OpenDocumentSpreadsheet
from odf.style import Style, TableColumnProperties, TableRowProperties, TextProperties
from odf.table import Table, TableColumn, TableRow, TableCell
from odf.text import P

def get_installed_software():
    uninstall_keys = [
        r"SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall",
        r"SOFTWARE\WOW6432Node\Microsoft\Windows\CurrentVersion\Uninstall"
    ]

    software_list = []

    for uninstall_key in uninstall_keys:
        try:
            reg_key = winreg.OpenKey(winreg.HKEY_LOCAL_MACHINE, uninstall_key)
        except FileNotFoundError:
            continue

        for i in range(0, winreg.QueryInfoKey(reg_key)[0]):
            try:
                sub_key_name = winreg.EnumKey(reg_key, i)
                sub_key = winreg.OpenKey(reg_key, sub_key_name)
                software = {}

                try:
                    software["name"] = winreg.QueryValueEx(sub_key, "DisplayName")[0]
                except FileNotFoundError:
                    continue

                try:
                    software["Version"] = winreg.QueryValueEx(sub_key, "DisplayVersion")[0]
                except FileNotFoundError:
                    software["Version"] = "Unknown"

                try:
                    software["Publisher"] = winreg.QueryValueEx(sub_key, "Publisher")[0]
                except FileNotFoundError:
                    software["Publisher"] = "Unknown"

                try:
                    software["Language"] = winreg.QueryValueEx(sub_key, "Language")[0]
                except FileNotFoundError:
                    software["Language"] = "Unknown"

                try:
                    software["Product code"] = winreg.QueryValueEx(sub_key, "ProductID")[0]
                except FileNotFoundError:
                    software["Product code"] = "Unknown"

                if "WOW6432Node" in uninstall_key:
                    software["Bits"] = "32"
                else:
                    software["Bits"] = "64" if platform.machine().endswith("64") else "32"

                # Extra kolom: Type programma (dummy voor nu)
                software["Type"] = "Unknown"

                software_list.append(software)

            except Exception:
                continue

    return software_list


def export_to_ods(software_list, filename="software-list.ods", name=""):
    doc = OpenDocumentSpreadsheet()

    # make a spreatcheat
    table = Table(name="Programmes on my laptop")

    colstyle = Style(name="colwidth", family="table-column")
    colstyle.addElement(TableColumnProperties(columnwidth="2.5cm"))
    doc.styles.addElement(colstyle)

    # add 8 colloms
    for _ in range(8):
        table.addElement(TableColumn(stylename=colstyle))

    # Row 2
    row = TableRow()
    cell = TableCell(numbercolumnsrepeated=6)
    cell.addElement(P(text="Programmes on my laptop"))
    row.addElement(cell)

    # cel with name
    cell = TableCell()
    cell.addElement(P(text=f"Your name: {name}"))
    row.addElement(cell)

    # empty cel
    row.addElement(TableCell())
    table.addElement(row)

    # Row 3
    headers = ["name of the software", "Version", "Language", "Bits (32/64)", "Publisher", "Product code", "Type programma"]
    row = TableRow()
    for h in headers:
        cell = TableCell()
        cell.addElement(P(text=h))
        row.addElement(cell)
    table.addElement(row)

    # Data
    for sw in software_list:
        row = TableRow()
        for key in ["name", "Version", "Language", "Bits", "Publisher", "Product code", "Type"]:
            cell = TableCell()
            cell.addElement(P(text=str(sw.get(key, ""))))
            row.addElement(cell)
        table.addElement(row)

    doc.spreadsheet.addElement(table)
    doc.save(filename)
    print(f"File saved as: {filename}")


if __name__ == "__main__":
    programs = get_installed_software()
    export_to_ods(programs, "software-list.ods", name="[name]")
