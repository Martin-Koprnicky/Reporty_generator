import tableauserverclient as TSC
import pandas as pd
import asyncio
import datetime
import os
from pypdf import PdfWriter, PdfReader
import io
user_name = os.getlogin()
main_path = fr'C:\Users\chaloupeckyt\OneDrive - VLP\Dokumenty'


def export_one(export, server):
    wb = server.views.get_by_id(export['id'])
    server.views.populate_pdf(wb, export['params'])
    try:
        stream = io.BytesIO(initial_bytes=wb.pdf)
        read_pdf = PdfReader(stream)
        return read_pdf
    except TSC.ServerResponseError:
        raise 
    except Exception as error:
        print("Nastala chyba: ", error)
        raise
    

async def async_export(export, server):
    return await asyncio.to_thread(export_one, export, server)

class Report:
    def __init__(self, workbook_name: str, must_contains: list[str] = [''], 
                pdf_params = TSC.PDFRequestOptions(
                    page_type=TSC.PDFRequestOptions.PageType.Unspecified, 
                    orientation=TSC.PDFRequestOptions.Orientation.Portrait),
                one_file_name: str = None, split_names: list[str] = None):
        self.workbook_name = workbook_name
        for str in must_contains:
            if ('(' in str) or (')' in str):
                raise ValueError('Parametr "must_contains" nesmí obsahovat závorky: "(" nebo ")"')
        self.must_contains = must_contains
        self.pdf_params = pdf_params

        self.one_file_name = one_file_name
        self.split_names = split_names


    def make_ids(self, server):
        all_workbooks = list(TSC.Pager(server.workbooks))
        wbook_names = [workbook.name for workbook in all_workbooks]
        wbook_ids = [workbook.id for workbook in all_workbooks]
        wbook_tags = [workbook.tags for workbook in all_workbooks]    
        w_df = pd.DataFrame({'id': wbook_ids, 
                            'name': wbook_names,
                            'tags': wbook_tags})
        
        wb_ids = w_df[(w_df['tags'].apply(lambda set: 'report' in set))&(w_df['name'] == self.workbook_name)].id.to_list()
        if len(wb_ids) == 1:
            w_id = wb_ids[0]
        else:
            if len(wb_ids) == 0:
                error_string = f"Nebyl vybrán žádný workbook. Má daný workbook {self.workbook_name} tag 'report'?"
            else:
                error_string = f"Nebyl vybrán jen jeden Workbook. Specificky jich bylo vybráno {len(wb_ids)}. Existuje tedy více reportů s názevm {self.workbook_name}"

            raise ValueError(error_string)

        all_views = list(TSC.Pager(server.views))
        view_names = [view._name for view in all_views]
        view_wbook_ids = [view._workbook_id for view in all_views]
        view_ids = [view.id for view in all_views]
        view_df = pd.DataFrame({'view_id': view_ids, 
                                'view_name': view_names, 
                                'wbook_id': view_wbook_ids})

        relevant_df = view_df[view_df['wbook_id']==w_id]
        podminka = pd.Series(True, index=relevant_df.index)
        for str in self.must_contains:
            podminka = podminka & relevant_df['view_name'].str.contains(str, regex = True)
        df = relevant_df[podminka][['view_name', 'view_id']].sort_values(by = ['view_name'], key = lambda col: col.str.lower())
        export_id = df['view_id'].to_list()
        if export_id == []:
            raise ValueError('Nebyl vybrát žadný Sheet.')
        export_list = [{'id': one,
                        'params': self.pdf_params} 
                    for one in export_id]
        return export_list

    async def gather_pdf(self, file_path):
        if self.split_names is None and self.one_file_name is None:
            raise ValueError('Soubor musí buď mít jedno jmeno nebo vícero jmen')
        if self.split_names is not None and self.one_file_name is not None:
            raise ValueError(f'Soubor nemůže mít jedno jmeno {self.one_file_name} a vícero jmen {self.split_names} současně')
        
        if self.split_names is None:
            if os.path.exists(file_path+'\\'+self.one_file_name):
                raise FileExistsError(f'Soubor {self.one_file_name} již existuje ve složce {file_path}')
        else:
            if all([os.path.exists(file_path+'\\'+name) for name in self.split_names]):
                raise FileExistsError(f'Soubory {', '.join(self.split_names)} již existují ve složce {file_path}')

        
        auth = TSC.PersonalAccessTokenAuth('AnalytikVLM2025',
                                           'Wya2UuOuRvyLckCudd/bvA==:hexR3IVMpil2LHIC6ZftmH8euEpFBCwp',
                                           site_id='vltavalabemedia')
        server = TSC.Server('https://eu-west-1a.online.tableau.com/', use_server_version=True)
        server.version = '3.22'
        with server.auth.sign_in(auth):
            export_pages = self.make_ids(server) 
            if self.split_names is not None:
                if len(self.split_names) != len(export_pages):
                    raise ValueError(f'Počet jmen souborů ({len(self.split_names)}) musí odpovídat počtu stránek pdf souboru ({len(export_pages)})')
            
            print('Začíná generování pdf souboru.')
            result_pdf = await asyncio.gather(*[async_export(one_page, server) for one_page in export_pages])
            print('Generovaní pdf souboru skončilo.\n')
            return result_pdf

    async def write_pdf(self, file_path):
        result_pdf = await self.gather_pdf(file_path)
        if self.split_names is not None:
            for pdf, name in zip(result_pdf, self.split_names):
                output = PdfWriter()
                output.add_page(pdf.pages[0])
                with open(file_path+'\\'+name, "wb") as f:
                    output.write(f)
        else:
            output = PdfWriter()
            for pdf in result_pdf: 
                output.add_page(pdf.pages[0])
            with open(file_path+'\\'+self.one_file_name, "wb") as file:
                output.write(file)

def file_names_list(one_report):
    if one_report.split_names == None:
        return [one_report.one_file_name]
    else:
        return one_report.split_names

class ReportGroup:
    def __init__(self, file_path, *group: Report):
        self.file_path = file_path
        self.group = group

    async def generate_pdf(self):
        print(f"Cesta k reportu je: {self.file_path}")
        for report in self.group:
            await report.write_pdf(self.file_path)

        
    def all_file_names(self, with_path = False):
        all_names = sum(list(map(file_names_list, self.group)), [])
        if with_path:
            all_names_with_names = [self.file_path+'\\'+name for name in all_names]
            return all_names, all_names_with_names
        else: 
            return all_names

class Datum:
    class Tyden:
        class Start:
            dt = datetime.date.today() - datetime.timedelta(days=datetime.date.today().weekday(), weeks=1)
            den = str(dt.day)
            tyden = str(dt.isocalendar()[1])
            rok = str(dt.year)
            mesic = str(dt.month)

            if len(tyden)==1:
                tyden = '0'+ tyden
            if len(mesic)==1:
                mesic = '0'+ mesic

        class End:
            dt = datetime.date.today()-datetime.timedelta(days=datetime.date.today().weekday()+1)
            den = str(dt.day)
            tyden = str(dt.isocalendar()[1])
            rok = str(dt.year)
            mesic = str(dt.month)

            if len(tyden)==1:
                tyden = '0'+ tyden
            if len(mesic)==1:
                mesic = '0'+ mesic

    class Mesic:
        dt = datetime.date.today()-datetime.timedelta(days=20)
        den = str(dt.day)
        tyden = str(dt.isocalendar()[1])
        rok = str(dt.year)
        mesic = str(dt.month)
        mesic_nazev =  {1: "Leden",
                        2: "Únor",
                        3: "Březen",
                        4: "Duben",
                        5: "Květen",
                        6: "Červen",
                        7: "Červenec",
                        8: "Srpen",
                        9: "Září",
                        10: "Říjen",
                        11: "Listopad",
                        12: "Prosinec"}[dt.month]

        if len(tyden)==1:
                tyden = '0'+ tyden
        if len(mesic)==1:
            mesic = '0'+ mesic


class Reporty:
    class Tydeni:
        Auto_magaziny = ReportGroup(
            main_path+fr'\01. Traffic and media\08. Auto-magazíny na denících & Tipcars.com\{Datum.Tyden.Start.rok}',
            Report(workbook_name = 'Auto–magazíny na denících & Tipcars.com – týdenní report',
                one_file_name = f"T{Datum.Tyden.Start.tyden}-{Datum.Tyden.Start.rok}-Auto-magazíny_na_denících_&_Tipcars.com.pdf")
        )

        Denik_online_performance = ReportGroup(
            main_path+fr'\01. Traffic and media\02. Denik on-line performance\{Datum.Tyden.Start.rok}',
            Report(workbook_name = 'Denik online performance - týdenní report',
                one_file_name = f"T{Datum.Tyden.Start.tyden}-{Datum.Tyden.Start.rok}-Deník_online_performance.pdf")
        )

        Magaziny_deniky = ReportGroup(
            main_path+fr'\01. Traffic and media\07. Magazíny na denících\{Datum.Tyden.Start.rok}',
            Report(workbook_name = 'Magazíny na denících – týdenní report',
                one_file_name = f"T{Datum.Tyden.Start.tyden}-{Datum.Tyden.Start.rok}-Magazíny_na_denících.pdf")
        )

        Online_performance = ReportGroup(
            main_path+fr'\01. Traffic and media\01. On-line publishing performance\{Datum.Tyden.Start.rok}',
            Report(workbook_name = 'Online publishing performance – týdenní report',
                one_file_name = f"T{Datum.Tyden.Start.tyden}-{Datum.Tyden.Start.rok}-Online_publishing_performance.pdf")
        )

        Top_20_claku_publishing = ReportGroup(
            main_path+fr'\01. Traffic and media\05. TOP 20 článků – online publishing weby\{Datum.Tyden.Start.rok}',
            Report(workbook_name='TOP 20 článků (řazení dle zdrojů) – týden',
                one_file_name=f"T{Datum.Tyden.Start.tyden}-{Datum.Tyden.Start.rok}-TOP_20_článků_online_publishing_weby_(pv_textu,_řazeno_dle_zdrojů).pdf"),
            Report(workbook_name = 'TOP 20 článků (pv 2) – týden',
                one_file_name=f"T{Datum.Tyden.Start.tyden}-{Datum.Tyden.Start.rok}-TOP_20_článků_online_publishing_weby_(pv_2).pdf"),
            Report(workbook_name = 'TOP 20 článků (pv textu) – týden', 
                one_file_name= f"T{Datum.Tyden.Start.tyden}-{Datum.Tyden.Start.rok}-TOP_20_článků_online_publishing_weby_(pv_textu).pdf"),
            Report(workbook_name = 'TOP 20 článků – týden',
                one_file_name= f"T{Datum.Tyden.Start.tyden}-{Datum.Tyden.Start.rok}-TOP_20_článků_online_publishing_weby_(pv).pdf")
        )

        Top_100_clanku_na_Deniku = ReportGroup(
            main_path+fr'\01. Traffic and media\11. Top 100 článků Deník\{Datum.Tyden.Start.rok}',
            Report(workbook_name = 'TOP 100 článků na Deníku (řazení dle Google aplikací)',
                must_contains=["01|02"],
                one_file_name=f"T{Datum.Tyden.Start.tyden}-{Datum.Tyden.Start.rok}-TOP_100_článků_Deníku_(Google_dicover).pdf")
        )

        Videa_videoreklamy = ReportGroup(
            main_path+fr'\01. Traffic and media\03. Video performance\{Datum.Tyden.Start.rok}',
            Report(workbook_name='Videa a videoreklamy',
                must_contains=["01|02"],
                one_file_name=f"T{Datum.Tyden.Start.tyden}-{Datum.Tyden.Start.rok}-Videa_a_videoreklamy.pdf")
        )

        Paywall = ReportGroup(
            main_path+fr'\02. Paywall\02.1 Paywall Deník\{Datum.Tyden.Start.rok}\Týden',
            Report(workbook_name = 'Paywall Deník – týdenní & měsíční report - PW 3.0 (skupiny předplatného)',
                must_contains = ["Týdenní report", "1|2|3|4|5"],
                one_file_name=f"T{Datum.Tyden.Start.tyden}-{Datum.Tyden.Start.rok}-Paywall-Deník.pdf")
        )
        
        HP_Seznam = ReportGroup(
            main_path+fr"\01. Traffic and media\04. Články na HP Seznam\{Datum.Tyden.Start.rok}\Týdny",
            Report(workbook_name = 'Články na HP Seznam',
                    must_contains = ["1"],
                    split_names=[f"T{Datum.Tyden.Start.tyden}-{Datum.Tyden.Start.rok}-Články_na_Seznam_HP-{name}.pdf" for name in  
                    ['Denik.cz',
                    'Dotyk',
                    'Dům_a_zahrada',
                    'iGlanc',
                    'iReceptář',
                    'Kafe',
                    'Kondice',
                    'Regiony',
                    'Vlasta']]),
            Report(workbook_name = 'Články na HP Seznam',
                must_contains = ["týdny", "Vývoj"],
                pdf_params=TSC.PDFRequestOptions(page_type=TSC.PDFRequestOptions.PageType.A4, 
                                                    orientation=TSC.PDFRequestOptions.Orientation.Landscape),
                one_file_name=f"T{Datum.Tyden.Start.tyden}-{Datum.Tyden.Start.rok}-Růstpokles_všech_webů_(s_průměrem_na_druhém_listu).pdf")
        )

        Top_podrubriky = ReportGroup(
            main_path+fr"\03. Redakce\06. TOP 10 článků z Podrubrik Deníku",
            Report(workbook_name='TOP Podrubriky Deníku',
                must_contains = ["týden"],
                pdf_params = TSC.PDFRequestOptions(page_type=TSC.PDFRequestOptions.PageType.Unspecified, 
                                    orientation=TSC.PDFRequestOptions.Orientation.Portrait),
                one_file_name = f"T{Datum.Tyden.Start.tyden}-{Datum.Tyden.Start.rok}-TOP 10_článků_z_Podrubrik_Deníku.pdf")
        )

        Sport_vykon_produkce = ReportGroup(
            main_path+fr'\03. Redakce\01. Sportovní redaktoři - výkon a produkce\{Datum.Tyden.Start.rok}',
            Report(workbook_name = 'Rubriky Deníku - výkon a produkce',
                must_contains = ["6.1|6.2"],
                one_file_name = f"T{Datum.Tyden.Start.tyden}-{Datum.Tyden.Start.rok}-Sportovní_redaktoři-výkon_a_produkce.pdf")
        ) 

        
    class Mesicni:
        Auto_magaziny = ReportGroup(
            main_path+fr'\01. Traffic and media\08. Auto-magazíny na denících & Tipcars.com\{Datum.Mesic.rok}',
            Report(workbook_name = 'Auto–magazíny na denících & Tipcars.com - měsíční report',
                one_file_name = f"M{Datum.Mesic.mesic}-{Datum.Mesic.rok}-Auto-magazíny_na_denících_&_Tipcars.com.pdf")
        )

        Denik_online_performance = ReportGroup(
            main_path+fr'\01. Traffic and media\02. Denik on-line performance\{Datum.Mesic.rok}',
            Report(workbook_name = 'Denik online performance - měsíční report',
                one_file_name = f"M{Datum.Mesic.mesic}-{Datum.Mesic.rok}-Deník_online_performance.pdf")
        )

        Magaziny_deniky = ReportGroup(
            main_path+fr'\01. Traffic and media\07. Magazíny na denících\{Datum.Mesic.rok}',
            Report(workbook_name = 'Magazíny na denících – měsíční report',
                one_file_name = f"M{Datum.Mesic.mesic}-{Datum.Mesic.rok}-Magazíny_na_denících.pdf")
        )

        Online_performance = ReportGroup(
            main_path+fr'\01. Traffic and media\01. On-line publishing performance\{Datum.Mesic.rok}',
            Report(workbook_name = 'Online publishing performance – měsíční report',
                one_file_name = f"M{Datum.Mesic.mesic}-{Datum.Mesic.rok}-Online_publishing_performance.pdf")
        )

        Top_20_claku_publishing = ReportGroup(
            main_path+fr'\01. Traffic and media\05. TOP 20 článků – online publishing weby\{Datum.Mesic.rok}',
            Report(workbook_name='TOP 20 článků (řazení dle zdrojů) – měsíc',
                one_file_name=f"M{Datum.Mesic.mesic}-{Datum.Mesic.rok}-TOP_20_článků_online_publishing_weby_(pv_textu,_řazeno_dle_zdrojů).pdf"),
            Report(workbook_name = 'TOP 20 článků (pv 2) – měsíc',
                one_file_name=f"M{Datum.Mesic.mesic}-{Datum.Mesic.rok}-TOP_20_článků_online_publishing_weby_(pv_2).pdf"),
            Report(workbook_name = 'TOP 20 článků (pv textu) – měsíc', 
                one_file_name= f"M{Datum.Mesic.mesic}-{Datum.Mesic.rok}-TOP_20_článků_online_publishing_weby_(pv_textu).pdf"),
            Report(workbook_name = 'TOP 20 článků – měsíc',
                one_file_name= f"M{Datum.Mesic.mesic}-{Datum.Mesic.rok}-TOP_20_článků_online_publishing_weby_(pv).pdf")
        )

        Top_100_clanku_na_Deniku = ReportGroup(
            main_path+fr'\01. Traffic and media\11. Top 100 článků Deník\{Datum.Mesic.rok}',
            Report(workbook_name = 'TOP 100 článků na Deníku (řazení dle Google aplikací)',
                must_contains=["03|04"],
                one_file_name=f"M{Datum.Mesic.mesic}-{Datum.Mesic.rok}-TOP_100_článků_Deníku_(Google_dicover).pdf")
        )

        Videa_videoreklamy = ReportGroup(
            main_path+fr'\01. Traffic and media\03. Video performance\{Datum.Mesic.rok}',
            Report(workbook_name='Videa a videoreklamy',
                must_contains=["03|04"],
                one_file_name=f"M{Datum.Mesic.mesic}-{Datum.Mesic.rok}-Videa_a_videoreklamy.pdf")
        )

        Facebook_Infotainment = ReportGroup(
            main_path+fr'\01. Traffic and media\10. Facebook a sociální sítě\{Datum.Mesic.rok}',
            Report(workbook_name='Facebook report - Infotainment',
                   must_contains=["Přehled nad všemi weby"],
                   one_file_name=f'M{Datum.Mesic.mesic}-{Datum.Mesic.rok}_Facebook_report_Infotainment.pdf')
        )
 
        Facebook_Denik = ReportGroup(
            main_path+fr'\01. Traffic and media\10. Facebook a sociální sítě\{Datum.Mesic.rok}',
            Report(workbook_name='Facebook report - Deník',
                   must_contains=['1|2|3|4|5|6|7|8|9|10|11'],
                   one_file_name=f'M{Datum.Mesic.mesic}-{Datum.Mesic.rok}_Facebook_report_Denik.pdf')
        )

        Netmonitor = ReportGroup(
            main_path+fr'\04. Mediální data\01. Netmonitor - porovnání s konkurencí (aktuální platný report)\{Datum.Mesic.rok}',
            Report(workbook_name = 'NetMonitor - porovnání s konkurencí',
                   must_contains = ['Celkový přehled|Porovnání s konkurencí'],
                   one_file_name = f'M{Datum.Mesic.mesic}-{Datum.Mesic.rok}-NetMonitor_report.pdf'),
            Report(workbook_name = 'NetMonitor - porovnání s konkurencí',
                   must_contains = ['VLM Packy'],
                   one_file_name = f'M{Datum.Mesic.mesic}-{Datum.Mesic.rok}-NetMonitor_Packy.pdf')       
        )

        Paywall = ReportGroup(
            main_path+fr'\02. Paywall\02.1 Paywall Deník\{Datum.Mesic.rok}\Měsíc',
            Report(workbook_name = 'Paywall Deník – týdenní & měsíční report - PW 3.0 (skupiny předplatného)',
                must_contains = ["Měsíční report", "1|2|3|4|5"],
                one_file_name=f"M{Datum.Mesic.mesic}-{Datum.Mesic.rok}-Paywall-Deník.pdf")
        )

        Reglog = ReportGroup(
            main_path+fr'\06. Čtenáři\09.1 Reglog\{Datum.Mesic.rok}',
            Report(workbook_name='REGLOG Statistika uživatelů I', one_file_name=f"M{Datum.Mesic.mesic}-{Datum.Mesic.rok}-Reglog_Statistiky_uživatelů.pdf")
        )

        Denik_User_Performance = ReportGroup(
            main_path+fr'\06. Čtenáři\09.2 Deník Users Performance\{Datum.Mesic.rok}',
            Report(workbook_name='Deník User Performance', one_file_name=f"M{Datum.Mesic.mesic}-{Datum.Mesic.rok}-Deník_Users_Performance.pdf")
        )


        Sociodemo = ReportGroup(
            main_path+fr'\04. Mediální data\02. Sociodemo\{Datum.Mesic.rok}',
            *[Report(workbook_name='Sociodemo report', must_contains=['SOCIODEMO'],
                   pdf_params=TSC.PDFRequestOptions(
                       page_type=TSC.PDFRequestOptions.PageType.Unspecified, 
                       orientation=TSC.PDFRequestOptions.Orientation.Portrait)
                       .vf('MY(date_from)', f'{Datum.Mesic.rok}{Datum.Mesic.mesic}')
                       .vf('node_name', "Internet," + filtr),        
                    one_file_name=f"M{Datum.Mesic.mesic}-{Datum.Mesic.rok}-SocioDemo-{web}.pdf")
                for filtr, web in zip(['VLTAVA-LABE-MEDIA\\, a.s. | annonce.cz',
                                        'VLTAVA-LABE-MEDIA\\, a.s. | cars.cz',
                                        'VLTAVA-LABE-MEDIA\\, a.s. | national-geographic.cz',
                                        'VLTAVA-LABE-MEDIA\\, a.s. | denik.cz',
                                        'VLTAVA-LABE-MEDIA\\, a.s. | dotyk.cz',
                                        'VLTAVA-LABE-MEDIA\\, a.s. | dumazahrada.cz',
                                        'VLTAVA-LABE-MEDIA\\, a.s. | iglanc.cz',
                                        'VLTAVA-LABE-MEDIA\\, a.s. | igurmet.cz',
                                        'VLTAVA-LABE-MEDIA\\, a.s. | ireceptar.cz',
                                        'VLTAVA-LABE-MEDIA\\, a.s. | kafe.cz',
                                        'VLTAVA-LABE-MEDIA\\, a.s. | kondice.cz',
                                        'VLTAVA-LABE-MEDIA\\, a.s. | poznatsvet.cz',
                                        'VLTAVA-LABE-MEDIA\\, a.s. | tipcars.com',
                                        'VLTAVA-LABE-MEDIA\\, a.s. | vlasta.cz']  ,
                                      ['Annonce',
                                        'Cars',
                                        'National_geographic',
                                        'Deník',
                                        'Dotyk',
                                        'Dům_a_zahrada',
                                        'iGlanc',
                                        'iGurmet',
                                        'iReceptář',
                                        'Kafe',
                                        'Kondice',
                                        'Poznatsvět',
                                        'Tipcars',
                                        'Vlasta'])])

        HP_Seznam = ReportGroup(
            main_path+fr'\01. Traffic and media\04. Články na HP Seznam\{Datum.Mesic.rok}\Měsíce',
            *[Report(workbook_name='Články na HP Seznam', must_contains = ['Články na Seznam HP'],
                   pdf_params=TSC.PDFRequestOptions(
                       page_type=TSC.PDFRequestOptions.PageType.Unspecified, 
                       orientation=TSC.PDFRequestOptions.Orientation.Portrait).vf('MY(Datum čtení měsíc)', f'{Datum.Mesic.rok}{Datum.Mesic.mesic}').vf('Rozdělení webů', filtr),
                    one_file_name=f"M{Datum.Mesic.mesic}-{Datum.Mesic.rok}-Články_na_Seznam_HP-{web}.pdf")
            for filtr, web in zip( 
                ['celostát', 'Dotyk', 'Dům a zahrada', 'iGlanc', 'iReceptář', 'Kafe', 'Kondice', 'regiony', 'Vlasta'],
                ['Denik.cz', 'Dotyk', 'Dům_a_zahrada', 'iGlanc', 'iReceptář', 'Kafe', 'Kondice', 'Regiony', 'Vlasta'])],
            Report(workbook_name='Články na HP Seznam', must_contains=['Vývoj, růst a podíl na celkovém trafficu'],
                   pdf_params = TSC.PDFRequestOptions(
                        page_type=TSC.PDFRequestOptions.PageType.Unspecified, 
                        orientation=TSC.PDFRequestOptions.Orientation.Landscape),
                    one_file_name=f"M{Datum.Mesic.mesic}-{Datum.Mesic.rok}-Články_na_Seznam_HP-Růstpokles_všech_webů.pdf")
        )