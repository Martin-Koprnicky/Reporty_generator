import pandas as pd
import asyncio
import datetime
import os
from pypdf import PdfWriter, PdfReader
import io
import tableauserverclient as TSC
user_name = os.getlogin()

dt = datetime.date.today() - datetime.timedelta(days=7)
tyden = str(dt.isocalendar()[1])
rok = str(dt.year)
mesic = str(dt.month)
if len(tyden)==1:
    tyden = '0'+tyden
if len(mesic)==1:
    mesic = '0'+mesic


def make_ids(worbook_name: str, must_contains: list, server, pdf_params):
    all_workbooks = list(TSC.Pager(server.workbooks))
    wbook_names = [workbook.name for workbook in all_workbooks]
    wbook_ids = [workbook.id for workbook in all_workbooks]
    w_df = pd.DataFrame({'id': wbook_ids, 'name': wbook_names})
    w_id = w_df[w_df['name'] == worbook_name].id.to_list()[0]

    all_views = list(TSC.Pager(server.views))
    view_names = [view._name for view in all_views]
    view_wbook_ids = [view._workbook_id for view in all_views]
    view_ids = [view.id for view in all_views]
    view_df = pd.DataFrame({'view_id': view_ids, 'view_name': view_names, 'wbook_id': view_wbook_ids})

    relevant_df = view_df[view_df['wbook_id']==w_id].copy()
    podminka = pd.Series(True, index=relevant_df.index)
    for str in must_contains:
        podminka = podminka & relevant_df['view_name'].str.contains(str)
    df = relevant_df[podminka][['view_name', 'view_id']].sort_values(by = ['view_name'], key = lambda col: col.str.lower())
    export_id = df['view_id'].to_list()
    export_list = [{'id': one, 'server': server, 'params': pdf_params} for one in export_id]
    return export_list

def export_one(export):
    wb = export['server'].views.get_by_id(export['id'])
    export['server'].views.populate_pdf(wb, export['params'])
    stream = io.BytesIO(initial_bytes=wb.pdf)
    read_pdf = PdfReader(stream)
    return read_pdf

async def async_export(export):
    return await asyncio.to_thread(export_one, export)
    
async def write_pdf(export_pages, path, name):
    result = await asyncio.gather(*[async_export(export_page) for export_page in export_pages])
    
    output = PdfWriter()
    for pdf in result: output.add_page(pdf.pages[0])
    with open(path+'\\'+name, "wb") as file:
        output.write(file)