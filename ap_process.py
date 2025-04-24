from PyQt5.QtWidgets import (QVBoxLayout, QHBoxLayout, QPushButton, QLabel,
                           QFileDialog, QTextEdit, QMessageBox, QListWidget)
from PyQt5.QtCore import Qt, QThread, pyqtSignal
from datetime import datetime, timedelta
import os
import win32com.client
from retro_style import RetroWindow, create_retro_central_widget
import pythoncom

def excel_to_df(input_path):
    """Convert Excel to DataFrame via CSV using isolated Excel application"""
    import pandas as pd


    excel = None
    wb = None
    temp_csv_path = None
    try:
        print(f"\nDetailed debug for reading {os.path.basename(input_path)}:")

        # Initialize COM for this thread
        pythoncom.CoInitialize()

        # Create new Excel instance
        print("1. Creating Excel application...")
        excel = win32com.client.DispatchEx("Excel.Application")
        excel.Visible = False
        excel.DisplayAlerts = False
        excel.EnableEvents = False
        excel.Interactive = False
        print("Excel application created")

        abs_input_path = os.path.abspath(input_path)
        temp_csv_path = abs_input_path.rsplit('.', 1)[0] + '_temp.csv'

        print("3. Opening workbook...")
        wb = excel.Workbooks.Open(abs_input_path)
        print("4. Saving as CSV...")
        wb.SaveAs(temp_csv_path, FileFormat=6)
        wb.Close(SaveChanges=False)
        wb = None

        print("6. Reading CSV file...")
        # Read CSV without doing any column mapping yet
        df = pd.read_csv(temp_csv_path, dtype={'Vendor Account Number': str, 'Vendor Routing Number': str})

        try:
            os.remove(temp_csv_path)
        except Exception as e:
            print(f"Warning: Could not delete temp file: {str(e)}")

        # Clean up the DataFrame
        df = df.replace('', pd.NA)
        df = df.dropna(how='all')

        # Only map columns if they exist and haven't been mapped yet
        if 'Due Date' in df.columns and 'Payment Date' not in df.columns:
            df = df.rename(columns={'Due Date': 'Payment Date'})
        if 'Amount Due' in df.columns and 'Pay $' not in df.columns:
            df = df.rename(columns={'Amount Due': 'Pay $'})

        return df

    except Exception as e:
        print(f"Error processing file: {str(e)}")
        return None

    finally:
        # Clean up in specific order
        if wb is not None:
            try:
                wb.Close(SaveChanges=False)
            except:
                pass
        if excel is not None:
            try:
                excel.Quit()
                del excel
            except:
                pass
        # Uninitialize COM
        pythoncom.CoUninitialize()




def clean_location_name(location):
    name = location.lower()
    remove_terms = ['llc', 'll', 'carrot', 'love', 'operating', 'operatin',]
    for term in remove_terms:
        name = name.replace(term, '')
    name = ' '.join(word.capitalize() for word in name.split())
    name = name.strip()
    name = name.replace(' ', '_')
    return name

def load_location_mapping():
    location_mapping = {
        ""
        "Carrot Leadership LLC": ["30000480952", "0952 CR CCD/PPD CARROT LEADERSHIP LLC"],
        "Carrot Express Franchise System LLC": ["30000481015", "1015 CR CCD/PPD CARROT EXPRESS FRANCHISE SYSTEM LL"],
        "Carrot Global LLC": ["30000482356", "2356 CR CCD/PPD CARROT GLOBAL LLC"],
        "Carrot Express Commissary LLC": ["30000488431", "8431 CR CCD/PPD CARROT EXPRESS COMMISSARY LLC"],
        "NY - Carrot Express Commissary": ["30000488431", "8431 CR CCD/PPD CARROT EXPRESS COMMISSARY LLC"],
        "Carrot Coral GablesLove LLC (Coral Gabes)": ["30000481123", "1123 CR CCD/PPD  CARROT LOVE LLC"],
        "Carrot Aventura Love LLC (Aventura)": ["30000481123", "1123 CR CCD/PPD  CARROT LOVE LLC"],
        "Carrot North Beach Love LL (North Beach)": ["30000481123", "1123 CR CCD/PPD  CARROT LOVE LLC"],
        "Carrot Downtown Love Two LLC": ["30000481258", "1258 CR CCD/PPD CARROT LOVE TWO LLC"],
        "Carrot Love City Place Doral Operating LLC": ["30000481978", "1978 CR CCD/PPD CARROT LOVE CITYPLACE DORAL OPERAT"],
        "Carrot Love Palmetto Park Operating LLC": ["30000482122", "2122 CR CCD/PPD CARROT LOVE PALMETTO PARK OPERATIN"],
        "Carrot Love Brickell Operating LLC": ["30000482104", "2104 CR CCD/PPD CARROT LOVE BRICKELL OPERATING LLC"],
        "Carrot Love West Boca Operating LLC": ["30000482140", "2140 CR CCD/PPD CARROT LOVE WEST BOCA OPERATING LL"],
        "Carrot Love Aventura Mall Operating LLC": ["30000482023", "2023 CR CCD/PPD CARROT LOVE AVENTURA MALL OPERATIN"],
        "Carrot Love Coconut Creek Operating LLC": ["30000482167", "2167 CR CCD/PPD CARROT LOVE COCONUT CREEK OPERATIN"],
        "Carrot Love Coconut Grove Operating LLC": ["30000482176", "2176 CR CCD/PPD CARROT LOVE COCONUT GROVE OPERATIN"],
        "Carrot Love Sunset Operating LLC": ["30000482212", "2212 CR CCD/PPD CARROT LOVE SUNSET OPERATING LLC"],
        "Carrot Love Pembroke Pines Operating LLC": ["30000594757", "4757 CR CCD/PPD CARROT LOVE PEMBROKE PINES OPERATI"],
        "Carrot Love Plantation Operating LLC": ["30000482149", "2149 CR CCD/PPD CARROT LOVE PLANTATION OPERATING L"],
        "Carrot Love River Lading Operating LLC": ["30000482230", "2230 CR CCD/PPD CARROT LOVE RIVER LANDING OPERATIN"],
        "Carrot Love Las Olas Operating LLC": ["30000482158", "2158 CR CCD/PPD CARROT LOVE LAS OLAS OPERATING LLC"],
        "Carrot Love Hollywood Operating LLC": ["30000482203", "2203 CR CCD/PPD CARROT LOVE HOLLYWOOD OPERATING LL"],
        "Carrot Sobe Love South Florida Operating C LLC": ["30000633502", "3502 CR CCD/PPD Carrot Love South Florida Operatin"],
        "Carrot Love South Florida Operating A LLC": ["30000633448", "2772 CR CCD/PPD Carrot Love South Florida Operatin"],
        "Carrot Flatiron Love Manhattan Operating LLC": ["30000482131", "2131 CR CCD/PPD Carrot Love Manhattan Operating LL"],
        "Carrot Love Bryant Park Operating LLC": ["30000482410", "2410 CR CCD/PPD CARROT LOVE BRYANT PARK OPERATING "],
        "Carrot Love 600 Lexington LLC": ["30000510616", "0616CR CCD/PPD CARROT LOVE LE"],
        "CARROT LOVE LIBERTY STREET LLC": ["30000674938", "4938 CR CCD/PPD CARROT LOVE LIBERTY STREET LLC"],
        "Carrot Holdings LLC": ["30000469729", "9729 CR CCD/PPD CARROT HOLDINGS LLC"],
        "Carrot Gem LLC": ["30000488503", "8503 CR CCD&PPD CARROT GEM LLC"],
        "Carrot Dream LLC": ["30000482266", "2226 CR CCD/PPD CARROT DREAM LLC"],
        "Carrot Love Dadeland Operating LLC": ["30000481834", "1834 CR CCD/PPD CARROT LOVE DADELAND OPERATING LLC"],
        "Beyond Branding LLC": ["30000566218", "6218 CR CCD/PPD BEYOND BRANDING LLC"]
    }
    return location_mapping


def load_vendor_mapping():
    # Special cases for vendor display names
    special_vendor_mapping = {
    "Action Plumbing and Heating Blackflow Corp": ["ACTION PLUMBING AND HEATING BACKFLO", "868616795", "21000021"],
    "Choice Mechanical Refrigeration Services": ["Choice Mechanical Refrigeration Ser", "229039391668", "63100277"],
    "Fire Zone Ventilation & Suppression Inc.": ["Fire Zone Ventilation & Suppression", "820586170", "21000021"],
    "Duke Martin Refrigeration & Air Cond Inc": ["Duke Martin Refrigeration & Air Con", "2000197483813", "63107513"],
    "Sunshine Cleaning Contractor & Services": ["Sunshine Cleaning Contractor & Serv", "287272279", "267084131"],
    "ALFRED I DUPONT BUILDING PARTNERSHIP LLP": ["ALFRED I DUPONT BUILDING PARTNERSHI", "30000355852", "66004367"],
    "Universal Environmental Consulting, Inc": ["Universal Environmental Consulting,", "97034953", "21411335"],
    "Hernan Gonzalez - Petit Cash SoFLC": ["Hernan Gonzalez Petit Cash SoFLC", "898146766009", "63100277"]
}

    # Regular vendor mapping (for demonstration, add all your vendors here)
    # "Vendor Name": ["Vendor Name", "Bank Account Number", "Routing Number"]
    vendor_mapping = {
        "Collins Fish & Seafood Inc": ["Collins Fish & Seafood Inc", "232388361", "267084131"],
        "International Delights LLC": ["International Delights LLC", "1501330058", "026013576"],
        "Williams Marble Polish Inc": ["Williams Marble Polish Inc", "961549268", "267084131"],
        "Firescan Alarms, Inc": ["Firescan Alarms, Inc", "3879659986", "267084131"],
        "Ana Sucre Petit Cash Manhattan": ["Ana Sucre Petit Cash Manhattan", "5937851391", "063107513"],
        "Ana Sucre Petit Cash Bryant Park": ["Ana Sucre Petit Cash Bryant Park", "5937851391", "063107513"],
        "Doris Araujo":["Doris Araujo", "826323610", "267078299"],
        "PeopleLinx": ["PeopleLinx", "40630224666040900", "121000248"],
        "Isaac Gabriel Holan Meza": ["Isaac Gabriel Holan Meza","219946940997", "101019644"],
        "Diony Alfonso Petit Cash Brickell": ["Diony Alfonso Petit Cash Brickell", "898149972515", "063100277"],
        "Cristhy Machin Petit Cash Downtown": ["Cristhy Machin Petit Cash Downtown", "898146486363","063100277"],
        "Baker305 LLC": ["Baker305 LLC", "898134329652", "063000047"],
        "BOCA Group International Inc": ["BOCA Group International Inc","257004036","021411335"],
        "Guardian Fire and Security, LLC": ["Guardian Fire and Security, LLC", "1503309722","026013576"],
        "Plantelier LLC": ["Plantelier LLC", "898099492695", "063100277"],
        "Arcane Coffee": ["Arcane Coffee", "10000251311106", "226082598"],
        "5A Healthy Restaurants LLC WKendall": ["5A Healthy Restaurants LLC WKendall", "898111217415", "063000047"],
        "5M Healthy Restaurants LLC Weston": ["5M Healthy Restaurants LLC Weston", "898119864349", "063100277"],
        "5AM Healthy Restaurants LLC Pinecre": ["5AM Healthy Restaurants LLC Pinecre", "898119862176", "063100277"],
        "Kimberly Hernandez Petit Cash Sobe": ["Kimberly Hernandez Petit Cash Sobe","229049169136", "063100277"],
        "UserWay INC": ["UserWay INC","9189439500", "026008866"],
        "LSI Industries Inc": ["LSI Industries Inc", "1004387606", "043000096" ],
        "Sy Electronics Corp": ["Sy Electronics Corp", "5761625960", "063107513"],
        "Patagonian Sea Products LLC": ["Patagonian Sea Products LLC", "227736359", "267084131"],
        "River Viiperi Inc": ["River Viiperi Inc","539265830", "322271627"],
        "Oscar Gastaudo PA.": ["Oscar Gastaudo PA.", "906969297", "21000021"],
        "Carrot Express Miami Shores LLC": ["Carrot Express Miami Shores LLC", "6766054644", "063107513"],
        "Green Planet Supplies LLC": ["Green Planet Supplies LLC", "1100022552072", "263191387"],
        "The new company CBPU LLC": ["The new company CBPU LLC", "898138986017", "63100277"],
        "Adriana Cribeiro Petit Cash MG": ["Adriana Cribeiro Petit Cash MG", "4443335454", "67014822"],
        "Gillman Consulting Inc": ["Gillman Consulting Inc", "656507370", "72000326"],
        "Isabel Arroyave": ["Isabel Arroyave", "229020770230", "63100277"],
        "Samuel Sultan": ["Samuel Sultan", "703926722", "267084131"],
        "Emporium Design": ["Emporium Design", "483087776835", "21000322"],
        "Forever Signs Inc": ["Forever Signs Inc", "4444269595", "67014822"],
        "FREEDOM SIGNS FLORIDA": ["FREEDOM SIGNS FLORIDA", "8100012672390", "263177903"],
        "Singer EVI LLC": ["Singer EVI LLC", "9856354049", "22000046"],
        "Elpo Electrical Contracting, Inc.": ["Elpo Electrical Contracting, Inc.", "313977529", "21000021"],
        "Abel Dominguez Petit Cash Hollywood": ["Abel Dominguez Petit Cash Hollywood", "898132513363", "63100277"],
        "Mauricio Romero": ["Mauricio Romero", "36195741002", "31176110"],
        "Fire Zone Services Inc": ["Fire Zone Services Inc", "4436307220", "26013673"],
        "Plumtech Services Inc": ["Plumtech Services Inc", "603827325", "267084131"],
        "Rachelle Azulay": ["Rachelle Azulay", "1566046023", "63107513"],
        "Keto KItchen 2GO": ["Keto KItchen 2GO", "9114740633", "266086554"],
        "Domaselo LLC": ["Domaselo LLC", "656944070636653", "121145349"],
        "Eny Diaz": ["Eny Diaz", "898141620865", "63100277"],
        "Claudia Parra Gabaldon": ["Claudia Parra Gabaldon", "3196758238", "67004764"],
        "Nixon Bracamontes - Elite Plumbers": ["Nixon Bracamontes - Elite Plumbers", "483100169604", "2000322"],
        "Douglas Guillen - Elite Plumbers": ["Douglas Guillen - Elite Plumbers", "590258292", "21000021"],
        "PeopleLinx": ["PeopleLinx", "40630224666040900", "121000248"],
        "B&H Photo Video Inc": ["B&H Photo Video Inc", "4125966952", "121000248"],
        "HCM Development Inc": ["HCM Development Inc", "767333987", "72000326"],
        "Felipe, Pedro": ["Felipe, Pedro", "612663602", "267084131"],
        "Bonilla Brenda": ["Bonilla Brenda", "587015796", "267084131"],
        "Lugo, Maria Fernanda": ["Lugo, Maria Fernanda", "4288726230", "67014822"],
        "Arias, Sabrina": ["Arias, Sabrina", "1100021787572", "263191387"],
        "Castano, Rosa": ["Castano, Rosa", "607996363", "267084131"],
        "Quintero, Osiel": ["Quintero, Osiel", "3128281726", "63107513"],
        "Uzcategui, Mariana": ["Uzcategui, Mariana", "898151074948", "63100277"],
        "Diaz, Gehovany": ["Diaz, Gehovany", "898153731483", "63100272"],
        "Brito, Daniela": ["Brito, Daniela", "898151337614", "63100277"],
        "Lopez, Jakelin": ["Lopez, Jakelin", "898151134697", "63100277"],
        "Faria Dias, Ramon": ["Faria Dias, Ramon", "6278931792", "63107513"],
        "Nieves Moreno, Kerwin": ["Nieves Moreno, Kerwin", "898151171333", "63100277"],
        "Barreto, Kelinyer": ["Barreto, Kelinyer", "898152314474", "63100277"],
        "Cabeza, Estefhani": ["Cabeza, Estefhani", "898135759917", "63100277"],
        "Alvarez, Abraham": ["Alvarez, Abraham", "483049543218", "21000322"],
        "Bentacourt, Laura": ["Bentacourt, Laura", "3866172673", "63107513"],
        "Reyna Linares, Angel A": ["Reyna Linares, Angel A", "483106402365", "21000322"],
        "Lopez, Viviana": ["Lopez, Viviana", "381069476381", "21200339"],
        "Counter Culture Coffee Inc": ["Counter Culture Coffee Inc", "4451348362", "111000012"],
        "Buckhead South Florida": ["Buckhead South Florida", "980080766", "124000054"],
        "Betancourt, Laura Contractor": ["Betancourt, Laura Contractor", "3866172673", "63107513"],
        "Leonard Brood": ["Leonard Brood", "898086382983", "63000047"],
        "Maria Fernanda Lugo PC Las Olas": ["Maria Fernanda Lugo PC Las Olas", "4288726230", "67014822"],
        "Laura Ortiz": ["Laura Ortiz", "898147214576", "63100277"],
        "JLCworks": ["JLCworks", "9118597756", "266086554"],
        "Venegas, Ruben": ["Venegas, Ruben", "898155308254", "63100277"],
        "Tovar, Adrian": ["Tovar, Adrian", "573836056", "21000021"],
        "Pablo Aguirre": ["Pablo Aguirre", "4444304226", "67014822"],
        "Restaurant City NJ": ["Restaurant City NJ", "1830588214", "21101108"],
        "Russell Film Company LLC": ["Russell Film Company LLC", "601039608", "267084131"],
        "Alexandra Sucre PC Commissary NY": ["Alexandra Sucre PC Commissary NY", "483096749219", "21000322"],
        "Sicifo solutions LLC": ["Sicifo solutions LLC", "571328023", "267084131"],
        "David Barreto": ["David Barreto", "898144024390", "63100277"],
        "Castro, Debora": ["Castro, Debora", "599856682", "21000021"],
        "LAM'S Snacks FL": ["LAM'S Snacks FL", "424080500", "267084131"],
        "Gables Miracle Mile LLC": ["Gables Miracle Mile LLC", "30000544384", "66004367"],
        "George Schkulnik": ["George Schkulnik", "781921819", "267084131"],
        "Brickell Owner LLC": ["Brickell Owner LLC", "4537339327", "121000248"],
        "UnclogMe LLC": ["UnclogMe LLC", "570650237", "267084131"],
        "PAN ON THE WAY LLC": ["PAN ON THE WAY LLC", "8050077408", "43000096"],
        "Mario Flores": ["Mario Flores", "1020000468058", "266080107"],
        "Edens Limited Partnership": ["Edens Limited Partnership", "1019291949", "43000096"],
        "Protano's Bakery LLC": ["Protano's Bakery LLC", "4407162921", "67005158"],
        "Carrot Express South Beach LLC": ["Carrot Express South Beach LLC", "338586321", "267084131"],
        "Edison Andrade": ["Edison Andrade", "4443846419", "67014822"],
        "Lam's Foods, Inc NYC": ["Lam's Foods, Inc NYC", "873877135", "21000021"],
        "One Mind Enterprises Inc.": ["One Mind Enterprises Inc.", "776562826", "267084131"],
        "United Restaurant Hood Services Corp": ["United Restaurant Hood Services Corp", "9294180055", "63107513"],
        "Samuel Sultan Morely LLC": ["Samuel Sultan Morely LLC", "926502185", "267084131"],
        "Paytronix Systems, Inc.": ["Paytronix Systems, Inc.", "3300422886", "121140399"],
        "Rubmary Delgado PC Boca East": ["Rubmary Delgado PC Boca East", "898130495988", "63100277"],
        "Julius Meinl North America LLC": ["Julius Meinl North America LLC", "1100020286383", "263191387"],
        "Hanlon Plumbing Co.": ["Hanlon Plumbing Co.", "8288854212", "63107513"],
        "Marisabel Graterol PC Dadeland": ["Marisabel Graterol PC Dadeland", "732178691", "267084131"],
        "Roach Buster Holding of America Inc": ["Roach Buster Holding of America Inc", "252612468805", "66011392"],
        "JLA Delivery Inc": ["JLA Delivery Inc", "102708986", "267084131"],
        "Universal Hood Tech, Inc": ["Universal Hood Tech, Inc", "10169350605", "66011392"],
        "MMG Sunset LLC": ["MMG Sunset LLC", "252575741505", "66011392"],
        "Recharte, Ramon": ["Recharte, Ramon", "1447801471", "63107513"],
        "Nick's Restaurant LLC": ["Nick's Restaurant LLC", "30000535420", "66004367"],
        "Refriconsa services": ["Refriconsa services", "7388121852", "63107513"],
        "Magda Lesmes Petit Cash AVE Mall": ["Magda Lesmes Petit Cash AVE Mall", "229058891369", "63100277"],
        "United of Omaha": ["United of Omaha", "148704077749", "104000029"],
        "Maria Laura Lugo": ["Maria Laura Lugo", "4288725159", "67014822"],
        "Argent Products, Corp": ["Argent Products, Corp", "3858893655", "267084131"],
        "United Hood Cleaning Corp.": ["United Hood Cleaning Corp.", "605973210", "21000021"],
        "Gillian Cruz Petit Cash Lexington": ["Gillian Cruz Petit Cash Lexington", "36049384605", "31176110"],
        "Ismael Noguera Petit Cash Mshores": ["Ismael Noguera Petit Cash Mshores", "898140539850", "63100277"],
        "Musa Products By Moroli USA INC": ["Musa Products By Moroli USA INC", "6695229473", "063107513"],
        "Guardian Fire and Security, LLC": ["Guardian Fire and Security, LLC", "1503309722", "26013576"],
        "LEX NY EQUITIES LLC": ["LEX NY EQUITIES LLC", "7028993389", "42000314"],
        "Hialeah Products CO.": ["Hialeah Products CO.", "1100022898106", "263191387"],
        "Nativo Acai": ["Nativo Acai", "1381403821", "63107513"],
        "Fabiola Cavalier PC Commissary": ["Fabiola Cavalier PC Commissary", "36246598153", "31176110"],
        "Dana Rozansky Consulting, LLC": ["Dana Rozansky Consulting, LLC", "898122342748", "63100277"],
        "Daniel Trillo": ["Daniel Trillo", "1424050456", "121000248"],
        "Carrot Express Midtown LLC": ["Carrot Express Midtown LLC", "690122067", "267084131"],
        "Security Fire Prevention Inc": ["Security Fire Prevention Inc", "40406358905", "66011392"],
        "Oriana Munoz Petit Cash Downton": ["Oriana Munoz Petit Cash Downton", "898132958054", "63100277"],
        "Gabriella Chehebar": ["Gabriella Chehebar", "153501050", "267084131"],
        "JC Electric Solutions Corp": ["JC Electric Solutions Corp", "520085229", "267084131"],
        "Laura Betancourt Petit Cash CCreek": ["Laura Betancourt Petit Cash CCreek", "3866172673", "63107513"],
        "Miami Prime Seafood": ["Miami Prime Seafood", "781981995", "267084131"],
        "Baldor Specialty Foods Inc": ["Baldor Specialty Foods Inc", "753975580", "21000021"],
        "Mario Laufer": ["Mario Laufer", "6767850404", "63107513"],
        "South Florida Paper Products LLC": ["South Florida Paper Products LLC", "2729195954", "63107513"],
        "Melon Corp DBA Melon Design Agency": ["Melon Corp DBA Melon Design Agency", "229057746392", "63100277"],
        "Ramon Dias Petit Cash Cgrove": ["Ramon Dias Petit Cash Cgrove", "6278931792", "63107513"],
        "Alexandra Sucre": ["Alexandra Sucre", "483096749219", "21000322"],
        "Karnis LLC": ["Karnis LLC", "898062092938", "63000047"],
        "Pablo V Maes Galindo": ["Pablo V Maes Galindo", "3107969384", "266086554"],
        "Ginette Salas PC River Landing": ["Ginette Salas PC River Landing", "898136570661", "63100277"],
        "515 LAS OLAS LLC": ["515 LAS OLAS LLC", "329681389006", "21300077"],
        "Mesa Plumbing": ["Mesa Plumbing", "898136779006", "63100277"],
        "Maria Vidal Petit Cash Plantation": ["Maria Vidal Petit Cash Plantation", "766130089", "267084131"],
        "David Casanova Petit Cash Brickell": ["David Casanova Petit Cash Brickell", "563931909", "267084131"],
        "Valeria Guzman Petit Cash Doral": ["Valeria Guzman Petit Cash Doral", "898143804841", "63100277"],
        "Angela Perreca": ["Angela Perreca", "6365868881", "63107513"],
        "Power Buddies Solutions LLC": ["Power Buddies Solutions LLC", "898150829864", "63100277"],
        "996826 Ontario Inc": ["996826 Ontario Inc", "1222089076", "267084199"],
        "Francisco Gutierrez Petit Cash Midtown": ["Francisco Gutierrez Petit Cash Midtown", "898130775569", "63100277"],
        "Negser Corp": ["Negser Corp", "918457664", "267084131"],
        "Raskin's Fish Market Inc.": ["Raskin's Fish Market Inc.", "2122423426", "21000322"],
        "MNO CREATIVE SOLUTIONS, LLC": ["MNO CREATIVE SOLUTIONS, LLC", "898023757199", "63100277"],
        "JOHRA W MULTISERVICE LLC": ["JOHRA W MULTISERVICE LLC", "6265671989", "63107513"],
        "Vicmarie Arevalo Petit Cash CoGA": ["Vicmarie Arevalo Petit Cash CoGA", "898122159326", "63100277"],
        "MSFM Corp": ["MSFM Corp", "1100026886137", "263191387"],
        "The Drinks Company": ["The Drinks Company", "939723737", "267084131"],
        "International Marketing": ["International Marketing", "4358451189", "26013673"],
        "David Lincoln Siegel": ["David Lincoln Siegel", "4539668357", "241070417"],
        "Sandra Gonzalez PC Manhattan": ["Sandra Gonzalez PC Manhattan", "483096145868", "21000322"],
        "Frank reza(gio) PC Bryant Park": ["Frank reza(gio) PC Bryant Park", "898132772890", "63100277"],
        "Felix Flemons": ["Felix Flemons", "5499850141", "63100277"],
        "Nicole Saraga": ["Nicole Saraga", "6541474877", "63107513"],
        "Rafael Zarante": ["Rafael Zarante", "898098959579", "63100277"],
        "Joshua Daniel Laufer": ["Joshua Daniel Laufer", "727651165", "267084131"],
        "Restaurant 365": ["Restaurant 365", "4577428758", "121000248"],
        "Elohim service and delivery": ["Elohim service and delivery", "898125154881", "63100277"],
        "Amad Construction LLC": ["Amad Construction LLC", "8535944238", "63107513"],
        "Suheily Briceño": ["Suheily Briceño", "898112611748", "63100277"],
        "Cesar Padron": ["Cesar Padron", "898103031599", "63100277"],
        "Evelyn Rojas": ["Evelyn Rojas", "229057938537", "63100277"],
        "Universal Environmental Consulting, Inc": ["Universal Environmental Consulting, Inc", "97034953", "21411335"],
        "Abraham Chehebar": ["Abraham Chehebar", "8902066028047", "44000804"],
        "Marcela Torres Petit Cash West Boca": ["Marcela Torres Petit Cash West Boca", "8619357299", "63107513"],
        "Alexandria Guerra": ["Alexandria Guerra", "662129763", "267084131"],
        "Paytronix Order & Delivery": ["Paytronix Order & Delivery", "3300422886", "121140399"],
        "Alejandra Bello Petit Cash NoBe": ["Alejandra Bello Petit Cash NoBe", "837823035", "267084131"],
        "Aventura Mall Venture": ["Aventura Mall Venture", "4980810899", "121000248"],
        "Ortus Engineering, P.A.": ["Ortus Engineering, P.A.", "6252365854", "21302567"],
        "Alberto Bassal": ["Alberto Bassal", "708199051", "267084131"],
        "Manuel Hackel": ["Manuel Hackel", "898091364682", "63100277"],
        "Panayoti Monpfeli": ["Panayoti Monpfeli", "898136632936", "63100277"],
        "Herlis Rico Petit Cash Sunset": ["Herlis Rico Petit Cash Sunset", "898127757312", "63100277"],
        "Jmeza Corp": ["Jmeza Corp", "1306714757", "63107513"],
        "Vanessa Rodriguez PC Hollywood": ["Vanessa Rodriguez PC Hollywood", "576088366", "267084131"],
        "Kellermeyer Bergensons Services LLC": ["Kellermeyer Bergensons Services LLC", "1453542697", "121000358"],
        "KBT Consulting LLC": ["KBT Consulting LLC", "624017409", "267084131"],
        "Park Square 5 LLC": ["Park Square 5 LLC", "239447871", "63104668"],
        "Studio Park LLC": ["Studio Park LLC", "1000050688", "26011701"],
        "Naval LLC": ["Naval LLC", "1020157826", "61100606"],
        "Zummo Inc.": ["Zummo Inc.", "509751", "66014069"],
        "Michael Schatten": ["Michael Schatten", "557190458", "267084131"],
        "Imperial Bag & Paper": ["Imperial Bag & Paper", "590412892", "21000021"],
        "Isabel Arroyave LLC": ["Isabel Arroyave LLC", "922011835", "267084131"],
        "Chefs 4 You LLC": ["Chefs 4 You LLC", "9135057161", "266086554"],
        "Cuickfix LLC": ["Cuickfix LLC", "898140473958", "63100277"],
        "Carrot Express Miami Shores LLC": ["Carrot Express Miami Shores LLC", "6766054644", "63107513"],
        "Barbara Fuenmayor Petit Cash PPines": ["Barbara Fuenmayor Petit Cash PPines", "3898250943", "63107513"],
        "Elisa Hernandez Petit Cash Las Olas": ["Elisa Hernandez Petit Cash Las Olas", "898135686844", "63100277"],
        "Ana C Sucre Sosa": ["Ana C Sucre Sosa", "5937851391", "63107513"],
        "Sun Biz Cable, LLC": ["Sun Biz Cable, LLC", "898117886947", "63100277"]

}


    # Combine both dictionaries
    vendor_mapping.update(special_vendor_mapping)
    return vendor_mapping

class APProcessThread(QThread):
    update_signal = pyqtSignal(str)
    finished_signal = pyqtSignal(bool, str, list)

    def __init__(self, input_files, output_dir):
        super().__init__()
        self.input_files = input_files
        self.output_dir = output_dir
        self.location_mapping = load_location_mapping()
        self.vendor_mapping = load_vendor_mapping()

    def run(self):
        import pandas as pd
        def log_message(msg):
            print(msg)
            self.update_signal.emit(str(msg))

        # Special locations that should be grouped together
        SPECIAL_LOCATIONS = {
            "Carrot North Beach Love LL (North Beach)",
            "Carrot Aventura Love LLC (Aventura)",
            "Carrot Coral GablesLove LLC (Coral Gabes)"
        }

        try:
            log_message("Starting AP payment processing...")
            today_date = (datetime.now() + timedelta(days=1)).strftime('%m-%d-%Y')

            # Create output directory with date
            output_dir = os.path.join(self.output_dir, f"ACHB {today_date}")
            os.makedirs(output_dir, exist_ok=True)

            total_files = len(self.input_files)
            processed_files = 0
            error_files = []
            unrecognized_vendors = set()
            for input_file in self.input_files:
                try:
                    log_message(f"\nProcessing file: {os.path.basename(input_file)}")

                    try:
                        if input_file.endswith('.xlsx'):
                            df = excel_to_df(input_file)
                            if df is None:
                                raise ValueError("Failed to read Excel file")
                        else:
                            df = pd.read_csv(input_file)

                        if df.empty:
                            raise ValueError("File contains no data")

                    except Exception as e:
                        log_message(f"File reading error: {str(e)}")
                        raise

                    log_message(f"Processing DataFrame with columns: {', '.join(df.columns)}")

                    # Group locations
                    unique_locations = df['Location'].unique()
                    special_locations_present = [loc for loc in unique_locations if loc in SPECIAL_LOCATIONS]
                    regular_locations = [loc for loc in unique_locations if loc not in SPECIAL_LOCATIONS]

                    # Process special locations together if any exist
                    if special_locations_present:
                        try:
                            log_message(f"Processing special locations together: {', '.join(special_locations_present)}")
                            location_data = df[df['Location'].isin(special_locations_present)]
                            # Add this right after creating location_data in both special and regular location processing:
                            print("Debug column data:")
                            print("Payment Date column:")
                            print(location_data['Payment Date'].values)
                            print("\nPay $ column:")
                            print(location_data['Pay $'].values)

                            try:
                                data = {
                                    'SEC Code': ['CCD'] * len(location_data),
                                    'Location Account Number': [self.location_mapping.get(location_data.iloc[0]['Location'], ['', ''])[0]] * len(location_data),
                                    'Location Subsidiary': [self.location_mapping.get(location_data.iloc[0]['Location'], ['', ''])[1]] * len(location_data)
                                }

                                vendors = []
                                account_numbers = []
                                routing_numbers = []

                                for vendor in location_data['Vendor']:
                                    vendor_details = self.vendor_mapping.get(vendor)
                                    if vendor_details is None:
                                        unrecognized_vendors.add(vendor)  # Add to set if not found
                                    vendor_details = self.vendor_mapping.get(vendor, [vendor, '', ''])
                                    vendors.append(str(vendor_details[0]))
                                    account_numbers.append(str(vendor_details[1]))
                                    routing_numbers.append(str(vendor_details[2]))

                                # Extract values as individual elements
                                data.update({
                                    'Vendor Display Name': vendors,
                                    'Vendor Account Number': account_numbers,
                                    'Vendor Routing Number': routing_numbers,
                                    'Inv. Date': location_data['Inv. Date'].values.tolist(),
                                    'Invoice': location_data['Invoice'].values.tolist(),
                                    'Payment Date': location_data['Payment Date'].values.tolist(),
                                    'Location': location_data['Location'].values.tolist(),
                                    'Pay $': location_data['Pay $'].values.tolist()
                                })

                                output_df = pd.DataFrame(data)
                                output_filename = f"ACHB_Carrot_Love_{today_date}.xlsx"
                                output_path = os.path.join(output_dir, output_filename)

                                log_message(f"Saving file: {output_filename}")
                                with pd.ExcelWriter(output_path, engine='openpyxl') as writer:
                                    output_df.to_excel(writer, index=False)
                                    worksheet = writer.sheets['Sheet1']
                                    for col in ['E', 'F']:
                                        for cell in worksheet[col]:
                                            cell.number_format = '@'

                            except Exception as e:
                                log_message(f"Error creating output file: {str(e)}")
                                raise

                        except Exception as e:
                            log_message(f"Error processing special locations: {str(e)}")
                            raise

                    # Process regular locations
                    for location in regular_locations:
                        try:
                            log_message(f"Processing location: {location}")
                            location_data = df[df['Location'] == location]
                            print("Debug column data:")
                            print("Payment Date column:")
                            print(location_data['Payment Date'].values)
                            print("\nPay $ column:")
                            print(location_data['Pay $'].values)

                            data = {
                                'SEC Code': ['CCD'] * len(location_data),
                                'Location Account Number': [self.location_mapping.get(location, ['', ''])[0]] * len(location_data),
                                'Location Subsidiary': [self.location_mapping.get(location, ['', ''])[1]] * len(location_data)
                            }

                            vendors = []
                            account_numbers = []
                            routing_numbers = []

                            for vendor in location_data['Vendor']:
                                vendor_details = self.vendor_mapping.get(vendor)
                                if vendor_details is None:
                                    unrecognized_vendors.add(vendor)  # Add to set if not found
                                vendor_details = self.vendor_mapping.get(vendor, [vendor, '', ''])
                                vendors.append(str(vendor_details[0]))
                                account_numbers.append(str(vendor_details[1]))
                                routing_numbers.append(str(vendor_details[2]))
                            # Extract values as individual elements
                            data.update({
                                'Vendor Display Name': vendors,
                                'Vendor Account Number': account_numbers,
                                'Vendor Routing Number': routing_numbers,
                                'Inv. Date': location_data['Inv. Date'].values.tolist(),
                                'Invoice': location_data['Invoice'].values.tolist(),
                                'Payment Date': location_data['Payment Date'].values.tolist(),
                                'Location': location_data['Location'].values.tolist(),
                                'Pay $': location_data['Pay $'].values.tolist()
                            })

                            output_df = pd.DataFrame(data)
                            clean_loc = clean_location_name(location)
                            output_filename = f"ACHB_{clean_loc}_{today_date}.xlsx"
                            output_path = os.path.join(output_dir, output_filename)

                            log_message(f"Saving file: {output_filename}")
                            with pd.ExcelWriter(output_path, engine='openpyxl') as writer:
                                output_df.to_excel(writer, index=False)
                                worksheet = writer.sheets['Sheet1']
                                for col in ['E', 'F']:
                                    for cell in worksheet[col]:
                                        cell.number_format = '@'

                        except Exception as e:
                            log_message(f"Error processing location {location}: {str(e)}")
                            raise

                    processed_files += 1
                    log_message(f"Successfully processed file {processed_files} of {total_files}")

                except Exception as e:
                    error_msg = f"Error processing {os.path.basename(input_file)}: {str(e)}"
                    log_message(error_msg)
                    error_files.append(error_msg)
                    continue

            if error_files:
                message = f"Processed {processed_files} of {total_files} files with {len(error_files)} errors:\n\n"
                message += "\n".join(error_files)
                self.finished_signal.emit(False, message, list(unrecognized_vendors))
            else:
                message = f"Successfully processed all {total_files} files!\nOutput location: {output_dir}"
                self.finished_signal.emit(True, message, list(unrecognized_vendors))

        except Exception as e:
            self.finished_signal.emit(False, f"An error occurred: {str(e)}")


def resource_path(relative_path):
    """Get absolute path to resource, works for dev and for PyInstaller"""
    import sys
    try:
        base_path = sys._MEIPASS
    except Exception:
        base_path = os.path.abspath(".")
    return os.path.join(base_path, relative_path)

class APWindow(RetroWindow):
    def __init__(self):
        super().__init__()
        self.initUI()
        self.selected_files = []


    def initUI(self):
        central_widget = create_retro_central_widget(self)
        layout = QVBoxLayout(central_widget)

        # Title
        title_label = QLabel("AP Payments", self)
        title_label.setAlignment(Qt.AlignCenter)
        title_label.setObjectName("title_label")
        layout.addWidget(title_label)

        # Instructions button
        self.instructions_button = QPushButton('Instructions')
        self.instructions_button.clicked.connect(self.show_instructions)
        layout.addWidget(self.instructions_button)

        # Input Files button and list
        input_layout = QVBoxLayout()
        self.input_button = QPushButton('Input Files')
        self.input_button.clicked.connect(self.select_files)
        input_layout.addWidget(self.input_button)
        self.file_list = QListWidget()
        self.file_list.setFixedHeight(150)  # Adjust this value as needed
        input_layout.addWidget(self.file_list)
        layout.addLayout(input_layout)

        # Run button
        self.run_button = QPushButton('RUN')
        self.run_button.clicked.connect(self.run_processing)
        layout.addWidget(self.run_button)

        # Console output
        self.console_output = QTextEdit()
        self.console_output.setReadOnly(True)
        layout.addWidget(self.console_output)

        self.setWindowTitle('AP Payments')
        self.resize(1000, 738)
        self.center()

    def center(self):
        from PyQt5.QtWidgets import QApplication
        frame_geometry = self.frameGeometry()
        center_point = QApplication.desktop().availableGeometry().center()
        frame_geometry.moveCenter(center_point)
        self.move(frame_geometry.topLeft())

    def select_files(self):
        files, _ = QFileDialog.getOpenFileNames(
            self, "Select Input Files", "", "Excel/CSV Files (*.xlsx *.csv)"
        )
        if files:
            self.selected_files.extend(files)
            self.update_file_list()
            self.console_output.append(f"Selected {len(files)} file(s)")

    def update_file_list(self):
        self.file_list.clear()
        for file in self.selected_files:
            self.file_list.addItem(os.path.basename(file))

    def select_output_directory(self):
        self.output_dir = QFileDialog.getExistingDirectory(self, "Select Output Directory")
        if self.output_dir:
            self.output_label.setText(f"Output directory: {self.output_dir}")
            self.console_output.append(f"Selected output directory: {self.output_dir}")

    def run_processing(self):
        if not self.selected_files:
            QMessageBox.warning(self, "Error", "Please select input files.")
            return

        # Get Downloads folder path
        downloads_path = os.path.join(os.path.expanduser('~'), 'Downloads')

        self.console_output.clear()
        self.console_output.append("Starting AP payment processing...")
        self.run_button.setEnabled(False)

        self.process_thread = APProcessThread(self.selected_files, downloads_path)
        self.process_thread.update_signal.connect(self.update_console)
        self.process_thread.finished_signal.connect(self.processing_finished)
        self.process_thread.start()

    def update_console(self, message):
        self.console_output.append(message)

    def processing_finished(self, success, message, unrecognized_vendors):
        self.run_button.setEnabled(True)

        # Build the complete message
        complete_message = message
        if unrecognized_vendors:
            complete_message += "\n\nUNRECOGNIZED VENDORS:\n"
            for vendor in unrecognized_vendors:
                complete_message += f"{vendor}\nFor now, put the vendor name manually in the bank. Code needs to update to include this vendor\n"

        # Show the appropriate dialog
        if success:
            QMessageBox.information(self, "Success", complete_message)
        else:
            QMessageBox.critical(self, "Error", complete_message)

        self.console_output.append(complete_message)




    def show_instructions(self):
        from PyQt5.QtWidgets import QScrollArea, QWidget, QDialog
        from PyQt5.QtGui import QPixmap
        instructions = """
 - For each company:

1. Select all invoices we will be paying.

2. "Save PMT Run" (DO NOT Create Payment yet).

3. Select Payment run from "Saved Payment Runs" (top right).

4. Download the Payment Run Excel file from R365.
"""
        # Create a custom dialog
        dialog = QDialog(self)
        dialog.setWindowTitle("Instructions")

        # Create main layout
        main_layout = QVBoxLayout(dialog)

        # Create scroll area to handle potential overflow
        scroll = QScrollArea()
        scroll.setWidgetResizable(True)
        scroll_widget = QWidget()
        scroll_layout = QVBoxLayout(scroll_widget)

        # Add text instructions
        text_label = QLabel(instructions)
        text_label.setWordWrap(True)
        scroll_layout.addWidget(text_label)

        # Add image
        image_label = QLabel()
        image_path = resource_path(os.path.join('assets', 'ap_instructions.png'))
        pixmap = QPixmap(image_path)
        scaled_pixmap = pixmap.scaled(750, 500, Qt.KeepAspectRatio, Qt.SmoothTransformation)
        image_label.setPixmap(scaled_pixmap)
        scroll_layout.addWidget(image_label)

        # Add remaining instructions
        remaining_text = """

5. In this App, click the "Input Files" button to select one or more Excel files containing AP data. You may select all files at once.

6. Click the "Output Directory" button to select where you want the processed files to be saved.

7. Click RUN to process the files.

8. Upload the achb file for one company at a time into CNB by selecting "UPLOAD FROM FILE" and choosing "ACHB" template.

9. Select payment date (tomorrow).

10. Initiate

Note: Make sure Excel files are not open when processing to avoid permission errors.
"""
        remaining_label = QLabel(remaining_text)
        remaining_label.setWordWrap(True)
        scroll_layout.addWidget(remaining_label)
        # Add second image
        second_image_label = QLabel()
        second_image_path = resource_path(os.path.join('assets', 'cnb_instructions.png'))  # Change filename here
        second_pixmap = QPixmap(second_image_path)
        second_scaled_pixmap = second_pixmap.scaled(750, 500, Qt.KeepAspectRatio, Qt.SmoothTransformation)
        second_image_label.setPixmap(second_scaled_pixmap)
        scroll_layout.addWidget(second_image_label)
        # Add some spacing at the bottom
        scroll_layout.addSpacing(10)
        # Set up scroll area
        scroll.setWidget(scroll_widget)
        main_layout.addWidget(scroll)
        # Create button layout
        button_layout = QHBoxLayout()
        ok_button = QPushButton("OK")
        ok_button.clicked.connect(dialog.accept)
        button_layout.addStretch()
        button_layout.addWidget(ok_button)
        # Add button layout to main layout
        main_layout.addLayout(button_layout)
        # Set dialog size
        dialog.setMinimumWidth(800)
        dialog.setMinimumHeight(500)
        dialog.exec_()
