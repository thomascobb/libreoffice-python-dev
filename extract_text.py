def macro(*args):
    # Ignore args as they will include the MouseEvent when this macro is run
    # from a button
    desktop = XSCRIPTCONTEXT.getDesktop()

    model = desktop.getCurrentComponent()

    # access the first sheet
    sheet = model.Sheets.getByIndex(0)

    # access cell C4
    cell1 = sheet.getCellRangeByName("C4")

    # set text inside
    cell1.String = "Hello world"


if "XSCRIPTCONTEXT" in globals():
    # Embedded
    g_exportedScripts = (macro,)
else:
    # Development
    import only_in_dev as XSCRIPTCONTEXT
    macro()
