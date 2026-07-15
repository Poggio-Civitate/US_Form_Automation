#Ai Mei Zhang
#Cole Adam Reilly
#Updated for 2026

from io import StringIO
from html.parser import HTMLParser
from pathlib import Path

#Solution taken from stack overflow:
#   https://stackoverflow.com/questions/753052/strip-html-from-strings-in-python
#   Answer by a user called Eloff, edited by a user called Olivier Le Floch
#   use to strip the html tags that are in the description object
class MLStripper(HTMLParser):
    def __init__(self):
        super().__init__()
        self.reset()
        self.strict = False
        self.convert_charrefs= True
        self.text = StringIO()
    def handle_data(self, d):
        self.text.write(d)
    def get_data(self):
        return self.text.getvalue()

#does the stripping, also calls helper function to get rid of the extra newlines that the tags needed, fixes some arrent spacing as well
def strip_tags(html):
    s = MLStripper()
    s.feed(html)
    return remove_extra_new_lines(s.get_data()).replace('    ',' ')
#Ai Mei designed function to find the randomly generated table names in the xlsx based on identifiable values
def tabnameViaHeader():
    for index in x.keys():
        df = x[index]
        for v in ['photos','catalogued','dateable','tabulated','plans','sample', 'sections']:
            if v in df.columns[0].lower():
                if v not in table_names:
                    table_names[v] = []
                table_names[v].append(f'{index}')
#must be an easier way to do this (regex)
#find a new line not after ending punc and remove it
#  needed as the form adds extra newlines to fit and looks bad in form
#  will fail to remove if the line naturally ends in punc, but that is uncommon occurence
#   could look for two newlines and remove all the ones that aren't in pairs perhaps?
def remove_extra_new_lines(s):
    ns = ''
    for i,x in enumerate(s):
        if x == '\n' and (s[i-1] not in '.?!'):
            continue
        else:
            ns += x
    return ns.replace('\n\n','\n')

def replace_if(var):
    if(var == 'Locus '):
        var = ''
    elif(',' in var):
        var = var.replace('Locus', 'Loci:')
    return var

from mailmerge import MailMerge #merges content into a docx template
import pandas as pd             #used to load excel files

#not used in code yet, but is used as reference
Area_names ={
    'ca': 'Civitate A',
    'cb': 'Civitate B',
    't': 'Tesoro'
    }
    #path to template file

template = "./US Form Template.docx"
template_eng = "./US Form Template (English).docx"
template_eng_sm = "./US Form Template (English) (smaller).docx"
    #path to file holding trench coordinates for the season, will need to be updated each year
Trench_Coordinates = './Trench_Coordinates.csv'

#excel file is one workbook with multiple pages, this is the best way I found to read them all
#   x is a dictionary of dataframes
xls = pd.ExcelFile('kobo_input_all.xlsx')
x = pd.read_excel(xls, sheet_name=None)
table_names = {}
tabnameViaHeader()
print(table_names)
#coord csv has a column for each trench, the header is the short name (Civitate A trench 95 becomes CA95)
#   the rows are written starting with NW going clockwise
#   coor_d{} is a dictionary lookup to speed up making multiple locus forms for the same trench
coorddf = pd.read_csv(Trench_Coordinates)
coor_d = {}

#load the main form sheet from the dict x, and replace all empty values with an empty string to legalize concat
#TODO FIND SUMMARY ENTRY WITHOUT YEAR
summarydf = x['Locus Summary Entry 2026']
summarydf = summarydf.fillna('')

    #for every locus that has a filled form
print(summarydf)
for index, row in summarydf.iterrows():
    
    #load the template at path and to populate at end
    document = MailMerge(template)
    document_eng = MailMerge(template_eng)
    document_eng_sm = MailMerge(template_eng_sm)

    #trench ids are saved like this in sheet, so break apart relevant info
    # {shortened}{Trench_numer}_{year of excavation}
    
    #TODO FIND TRENCH IDBY KEYWORD TRENCH
    id_break = row['Trench ID'].split('_')

    #sanity check, year of trench id matches year of current season
    #TODO FIND SEASON ENTRY BY KEYWORD SEASON
    assert(id_break[1] == str(row['Field Season']))

    #will need to be expanded to include other designations
    #could be made quicker/easier with the dictionary of names
    if id_break[0].lower().startswith('t') :
        area = "Tesoro"
        short = "T"
        trench_number = id_break[0][1:]
    elif id_break[0].lower().startswith('ca'):
        area = "Civitate A"
        short = "CA"
        trench_number = id_break[0][2:]
    elif id_break[0].lower().startswith('cb'):
        area = "Civitate B"
        short = "CB"
        trench_number = id_break[0][2:]
    else:
        area = "NOT YET IMPLEMENTED"
        short = "NA"
        trench_number = "ERROR"
        raise NotImplementedError

        #id is based on the order the forms are done, and are only used to relate other sheets 
            #entries back with the main sheet
    current_id = row['_index']
    #create paths to save file and folder structure
    t_name = short + trench_number
    output_path = f"./Output/{t_name}/ITA/"
    #TODO FIND LOCUS ID ENTRY BY KEYWORD LOCUS
    output_name = f"US Locus {row["Locus ID"]}.docx"
    output_path_eng = f"./Output/{t_name}/ENG/"
    output_name_eng = f"SU Locus {row["Locus ID"]} Form.docx"
    
    output_path_eng_sm = f"./Output/{t_name}/ENG/ENG_SM/"
    output_name_eng_sm = f"SU Locus {row["Locus ID"]} Form.docx"

    #if the shortened name isn't in the coordinate dictionary, add it will stripping out useless stuff
    ##print(coorddf)
    if t_name not in coor_d:
        coor_d[t_name] = '\n'.join(coorddf[t_name].dropna().tolist())
    coor = coor_d[t_name]

    
    Path(output_path).mkdir(parents=True, exist_ok=True)
    Path(output_path_eng).mkdir(parents=True, exist_ok=True)
    Path(output_path_eng_sm).mkdir(parents=True, exist_ok=True)
        
    #all sections are used to grab data from excel and reformat it for the US Form

    #natural/artificial check
    #TODO FIND Nat/Art ENTRY BY KEYWORD Artificial

    if 'natural' in row['Natural or Artificial'].lower():
        nat = 'X'
        art = ''
    elif 'artificial' in row['Natural or Artificial'].lower():
        nat = ''
        art = 'X'
    else:
        print("ERROR: Neither Artificial or Natural")
    #plan numbers
    plans = ''
    try:
        for j in table_names['plans']:
            for i,r in x[j].iterrows():
                if r['_parent_index'] == current_id:
                    if plans != '':
                        plans += '\n '
                    plans += r['Associated Plans']
    except KeyError as e:
        print("No Associated Plans")
        print("-1",e)
    if(plans == ''):
        plans = 'None'

    #grab all section names
    secs = ''
    try:
        for j in table_names['sections']:
            for i,r in x[j].iterrows():
                if r['_parent_index'] == current_id:
                    if secs != '':
                        secs += '\n '
                    secs += r['Associated Sections']
    except KeyError as e:
        print("No Associated Sections")
        print(e)
    #grab all photo names
    photos = ''
    try:
        for j in table_names['photos']:
            for i,r in x[j].iterrows():
                if r['_parent_index'] == current_id:
                    if photos != '':
                        photos += '\n'
                    photos += r['Associated Photos']
    except KeyError as e:
        print("No Associated Photos")
        print("-3",e)
    if photos == '':
        photos = 'None'
    #grab all special artifacts numbers, just need to be the number ie PC20230066 is just 66 
        #or sf-t103-2023-4-8 becomes 8
    specials = ''
    try:
        for j in table_names['catalogued']:
            for i,r in x[j].iterrows():
                if r['_parent_index'] == current_id:
                    if specials != '':
                        specials += ', '
                    if 'Contains Special Finds' in r and '-' in r['Contains Special Finds']:
                        lis = r['Contains Special Finds'].split('-')
                        specials += lis[-1]
                    else:
                        try:
                            specials += r['Contains Special Finds']
                        except: 
                            specials += r['Contains Catalogued Special Finds']
    except KeyError as e:
        print("No Special Finds")
        print("-2",e)
    if specials == '':
        specials = 'None'

    #tabulated counts structured {type}: {quantity} {measurement}\n  
    #                         ie: Tile: 13 Bowls\n
    tab = ''
    try:
        for j in table_names['tabulated']:
            for i,r in x[j].iterrows():
                if r['_parent_index'] == current_id:
                    new_str = ''
                    if tab != '':
                        new_str += '\n'
#TODO FIX THIS TO FIND THE RIGHT NAMES PROGRAMMATICALLY
                    if r['Tabulated Material Type'] == 'Other':
                        new_str += str(r['Tabulated Material Type (Other)'])
                    else:
                        new_str += str(r['Tabulated Material Type'])
                    new_str += ': ' + str(r['Tabulated Count']) + ' ' + str(r['Tabulated Count Type'])
                    if(not (new_str.startswith('nan: nan') or ' nan ' in new_str or 'nan: ' in new_str) and ': 0.0 ' not in new_str):
                        #if('Ceramic' in new_str ):
                        #    if(': 1.0 ' in new_str):
                        #        new_str = new_str.replace('Objects','Sherd')
                        #    else:
                        #        new_str = new_str.replace('Objects','Sherds')
                        #elif('Objects' in new_str):
                        #    if(': 1.0 ' in new_str):
                        #        new_str = new_str.replace('Objects','Fragment')
                        #    else:
                        #        new_str = new_str.replace('Objects','Fragments')
                        if('.0 ' in new_str):
                            new_str = new_str.replace('.0 ', ' ')
                        tab += new_str.replace('_',' ')
    except KeyError as e:
        print("No Tabulated Counts")
        print("-2",e)
    if tab=='':
        tab='None'

    #datable elements, simply grab whatever they input in the form comma sep
    datable = ''
    try:
        for j in table_names['dateable']:
            for i,r in x[j].iterrows():
                if r['_parent_index'] == current_id:
                    if datable != '':
                        datable += ', '
                    datable += str(r['Dateable Elements'])
    except KeyError as e:
        print("No Datable Elements")
        print("-1",e)
    if datable =='':
        datable = 'None'
        
        
    #period refers to the phase name (orientalizing, archaic)
    #TODO FIND Locus Period ENTRY BY KEYWORD Period?
    period = ''
    #period = ', '.join(row['Locus Periods'].split(' ')).replace('_', ' ').capitalize()
    for h in summarydf.columns.values.tolist():
        if "Locus Periods/" in h and row[h] != 0:
            if period != '':
                period += ', '
            period += h.split('/')[1].capitalize()
    #prelim phasing refers to the year range
    #TODO FIND Locus Period ENTRY BY KEYWORD Phasing?
    prelim = row['Preliminary Phasing'].replace('_', ' to ').capitalize()

    #Positioning/stratig
    #   all are the same structure comma seperated values
    #   try-excepts are used to as if in a season, no locus have a specific stratig type it will not appear in dicitonary, leading to a key error
    #       once one locus has one of that type the key error will stop
    #           if none have an anterior to locus code shouldn't crash 
    dic = {
            'same_as':
            {'stuff':'Locus ',
             'table':'group_strat_Same_As',
             'col':'Same As Locus'},
            'is_bound_to':
            {'stuff':'Locus ',
             'table':'group_strat_Is_Bound_To',
             'col':'Is Bound To Locus'},
        
            'ant':
            {'stuff':'Locus ',
             'table':'group_strat_Anterior_To',
             'col':'Anterior To Locus'} ,
            'post':
            {'stuff':'Locus ',
             'table':'group_strat_Posterior_To',
             'col':'Posterior To Locus'},
            
            'above':
            {'stuff':'Locus ',
             'table':'group_strat_Above',
             'col':'Above Locus'} ,
            'below':
            {'stuff':'Locus ',
             'table':'group_strat_Below',
             'col':'Below Locus'},
            
            'underlies':
            {'stuff':'Locus ',
             'table':'group_strat_Underlies',
             'col':'Underlies Locus'},
            'overlies':
            {'stuff':'Locus ',
             'table':'group_strat_Overlies',
             'col':'Overlies Locus'},
            
            'is_cut_by':
            {'stuff':'Locus ',
             'table':'group_strat_Is_Cut_By',
             'col':'Is Cut By Locus'},
            'cuts':
            {'stuff':'Locus ',
             'table':'group_strat_Cuts',
             'col':'Cuts Locus'},
           
            'is_filled_by':
            {'stuff':'Locus ',
             'table':'group_strat_Is_Filled_By',
             'col':'Is Filled By Locus'},
            'fills':{'stuff':'Locus ',
             'table':'group_strat_Fills',
             'col':'Fills Locus'}
           }

    for key in dic:
        strat_dc = dic[key]
        try:
            if strat_dc['table'] in x:
                for i,r in x[strat_dc['table']].iterrows():
                    if r['_parent_index'] == current_id:
                        if strat_dc['stuff'] != 'Locus ':
                            strat_dc['stuff'] += ', '
                        strat_dc['stuff'] += str(r[strat_dc['col']])
        except KeyError as p:
            print(p, 'is not found')


    #Change Locus to Loci, or make empty if none
    for k in dic:
        dic[k]['stuff'] = replace_if(dic[k]['stuff'])

    #Recovery techniques are lacking in detail in the Kobo form, so similarly the detail is low in US form
    #should be fine, a simple yes/no for the recovery should be okay
    #TODO FIND Rec Technique ENTRY BY KEYWORD ???
    siev = ''
    if str(row['Recovery Techniques/Dry Sieved']) == '1.0':
        siev += 'Dry Sieved'
    if str(row['Recovery Techniques/Wet Sieved']) == '1.0':
        if siev != 'Locus ':
            siev += ' and '
        siev += "Wet Sieved"
    if siev == '':
        siev = 'None'
    flotation = ''
    if str(row['Recovery Techniques/Flotation']) == '1.0':
        flotation += 'Floatation Samples Taken'
    if flotation == '':
        flotation = 'None'

    #list of what samples were taken
    samples = ''
    try:
        for j in table_names['sample']:
            for i,r in x[j].iterrows():
                if r['_parent_index'] == current_id:
                    if samples != '':
                        samples += ', '
                    samples += str(r['Contains Samples'])
    except KeyError as e:
        print(e, 10, "Key error")
        print("No Samples Taken?")
    finally:
        if samples == '':
            samples = 'None'
    #complex data has been grabbed, the rest can be done in line
        
        #TODO FIND all these guys [] ENTRY BY KEYWORD ???
    document.merge(
        Locus_Number = str(row['Locus ID']),
        Locus_Number_2 = str(row['Locus ID']),
        Year = str(row['Field Season']),
        Area = area,
        Trench_Number = trench_number,
        Trench_Coordinates = coor,
        Elevation_or_Depth = str(row['Depth (cm)']) + ' cm deep',
        Natural = nat,
        Artificial = art,
        Plan_Number = plans,
        Section_Number = secs,
        Photo_Names = photos,
        Catalogue_Numbers = specials,
        Definition_and_Position = str(row['Definition and Position']),
        Criteria_for_Distinction = str(row['Criteria for Distinction']),
        Mode_of_Formation = str(row['Mode of Formation']),
        Inorganic_List = str(row['Inorganic Components']),
        Organic_List = str(row['Organic Components']),
        Consistency = (str(row['Deposit Compaction']) + ' compaction').capitalize(),
        Color = str(row['Munsell Color']),
        Dimensions = str(row['Width (m)']) + ' m wide x ' + str(row['Length (m)']) + ' m long',
        State_of_Conservation = str(row['State of Conservation']),
        Description = strip_tags(str(row['General Description'])),
        Equal_To = dic['same_as']['stuff'],          
        Is_Bound_To = dic['is_bound_to']['stuff'],   
        Above = dic['above']['stuff'],            
        Below = dic['below']['stuff'],             
        Underlies = dic['underlies']['stuff'],       
        Overlies = dic['overlies']['stuff'],         
        Is_Cut_By = dic['is_cut_by']['stuff'],       
        Cuts = dic['cuts']['stuff'],                 
        Is_Filled_By = dic['is_filled_by']['stuff'], 
        Fills = dic['fills']['stuff'],               
        Previous_Loci = dic['post']['stuff'],        
        Following_Loci = dic['ant']['stuff'],        
        Observations = str(row['Observations']),
        Interpretations = str(row['Interpretation']),
        Datable_Elements = datable,
        Dating = prelim,
        Period_or_Phase = str(period),
        Quantity_Report = tab,
        Samples = samples,
        Flotation = flotation,
        Sieving = siev,
        Reliability = str(row['Stratigraphic Reliability']).capitalize()
    )
    document_eng.merge(
        Locus_Number = str(row['Locus ID']),
        Locus_Number_2 = str(row['Locus ID']),
        Year = str(row['Field Season']),
        Area = area,
        Trench_Number = trench_number,
        Trench_Coordinates = coor,
        Elevation_or_Depth = str(row['Depth (cm)']) + ' cm deep',
        Natural = nat,
        Artificial = art,
        Plan_Number = plans,
        Photo_Names = photos,
        Catalogue_Numbers = specials,
        Definition_and_Position = str(row['Definition and Position']),
        Criteria_for_Distinction = str(row['Criteria for Distinction']),
        Mode_of_Formation = str(row['Mode of Formation']),
        Inorganic_List = str(row['Inorganic Components']),
        Organic_List = str(row['Organic Components']),
        Consistency = (str(row['Deposit Compaction']) + ' compaction').capitalize(),
        Color = str(row['Munsell Color']),
        Dimensions = str(row['Width (m)']) + ' m wide x ' + str(row['Length (m)']) + ' m long',
        State_of_Conservation = str(row['State of Conservation']),
        Description = strip_tags(str(row['General Description'])),
        Equal_To = dic['same_as']['stuff'],          
        Is_Bound_To = dic['is_bound_to']['stuff'],   
        Above = dic['above']['stuff'],            
        Below = dic['below']['stuff'],             
        Underlies = dic['underlies']['stuff'],       
        Overlies = dic['overlies']['stuff'],         
        Is_Cut_By = dic['is_cut_by']['stuff'],       
        Cuts = dic['cuts']['stuff'],                 
        Is_Filled_By = dic['is_filled_by']['stuff'], 
        Fills = dic['fills']['stuff'],               
        Previous_Loci = dic['post']['stuff'],        
        Following_Loci = dic['ant']['stuff'],
        Observations = str(row['Observations']),
        Interpretations = str(row['Interpretation']),
        Datable_Elements = datable,
        Dating = prelim,
        Period_or_Phase = str(period),
        Quantity_Report = tab,
        Samples = samples,
        Flotation = flotation,
        Sieving = siev,
        Reliability = str(row['Stratigraphic Reliability']).capitalize()
    )
    document_eng_sm.merge(
        Locus_Number = str(row['Locus ID']),
        Locus_Number_2 = str(row['Locus ID']),
        Year = str(row['Field Season']),
        Area = area,
        Trench_Number = trench_number,
        Trench_Coordinates = coor,
        Elevation_or_Depth = str(row['Depth (cm)']) + ' cm deep',
        Natural = nat,
        Artificial = art,
        Plan_Number = plans,
        Photo_Names = photos,
        Catalogue_Numbers = specials,
        Definition_and_Position = str(row['Definition and Position']),
        Criteria_for_Distinction = str(row['Criteria for Distinction']),
        Mode_of_Formation = str(row['Mode of Formation']),
        Inorganic_List = str(row['Inorganic Components']),
        Organic_List = str(row['Organic Components']),
        Consistency = (str(row['Deposit Compaction']) + ' compaction').capitalize(),
        Color = str(row['Munsell Color']),
        Dimensions = str(row['Width (m)']) + ' m wide x ' + str(row['Length (m)']) + ' m long',
        State_of_Conservation = str(row['State of Conservation']),
        Description = strip_tags(str(row['General Description'])),
        Equal_To = dic['same_as']['stuff'],          
        Is_Bound_To = dic['is_bound_to']['stuff'],   
        Above = dic['above']['stuff'],            
        Below = dic['below']['stuff'],             
        Underlies = dic['underlies']['stuff'],       
        Overlies = dic['overlies']['stuff'],         
        Is_Cut_By = dic['is_cut_by']['stuff'],       
        Cuts = dic['cuts']['stuff'],                 
        Is_Filled_By = dic['is_filled_by']['stuff'], 
        Fills = dic['fills']['stuff'],               
        Previous_Loci = dic['post']['stuff'],        
        Following_Loci = dic['ant']['stuff'],
        Observations = str(row['Observations']),
        Interpretations = str(row['Interpretation']),
        Datable_Elements = datable,
        Dating = prelim,
        Period_or_Phase = str(period),
        Quantity_Report = tab,
        Samples = samples,
        Flotation = flotation,
        Sieving = siev,
        Reliability = str(row['Stratigraphic Reliability']).capitalize()
    )
    document.write(output_path + output_name)
    document_eng.write(output_path_eng + output_name_eng)
    document_eng_sm.write(output_path_eng_sm + output_name_eng_sm)