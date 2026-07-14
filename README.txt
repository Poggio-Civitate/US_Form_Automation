Poggio Civitate Archaeological Project, made in 2023
Author: Cole Adam Reilly 

The Problem:
    Trench Supervisors need to fill out two sets of forms when they close a Locus (The stratigraphic unit for this site). One is hosted on Kobotoolbox and integrates with our database OpenContext.org (called the kobo form). The other is a governmental form required by the soprintendente of archaeology for our region (called the US form).
    The kobo form works much like a google forms survey, is relatively user friendly and made by us using our sites language/terminology.
    The US form is a standard for the region, is written in italian, and made to accommodate the varing language of every site, it can be rather confusing to educate our excavators on the nuances of the form.
    The workflow for filling US forms is significantly slower, requiring you to duplicate all your information in distinct word files, one for each locus. This resulted in multiple hours of extra labor for each trench, with some have upwards of 30 distinct loci in any given trench, resulting in 30 kobo forms, and 30 distinct US Forms.

The Solution:
    Kate Kriendler (head of Excavation) approached myself and Eric Kansa (co-creator of our database) to find a way to automate the population of the US form. 
    Trench supervisors should fill out the kobo forms, and at the end of the season run a program to grabbed all the information and populate it into .docx files.
    
    To do this several things needed to happen:
        1. The kobo form was expanded to account for information only asked for in the US form. This was done by Eric and Kate.
        2. A US Form word document needed to be populated with mailmerge fields to allow for population. This was done by me, with the helpf of a guide to the US Form written by Ann Glennie (catalogour) and Kate.
        3. A python file needed to be written to read the xlsx output of the kobo form and populate the template US Form. This was done by me.

How to run the program?
    currently the solution is written as a .py file and requires the correct files/folder system as well as some presinstalled libraries. I might create a .ipynb file to better the system.

    run the file through cmd (terminal for mac) or through an IDE like VSCode open to the correct folder. 
        python compile_forms.py

    The file needs a couple things to run correctly, two libraries that must be installed beforehand, and two files that need to be named correctly and put in the correct space. Those are outlined below:

    Python Requirements: Download Python3 and the following libraries
        mailmerge
        pandas
    Should be installed using Pip, instructions to that can be found on the web.
    html.parser, os, and stringIO are all apart of the standard library and do not need to be installed.

    All locus forms should be downloaded from the kobo website: 
        https://kform.opencontext.org/#/forms
    navigate to the Locus Summary Entry form for this year
      click the data tab
         click the downloads tab
            select export type as xls, and 'XML values and headers'
            click export
            in exports section click download on generated export
            rename file to 'kobo_input_all.xlsx' and move the file to the same folder as compile_forms.py

    The opening trench coordinates need to be compiled into a CSV called 'Trench_Coordinates.csv'.
        Each column header should be the shortened designation for the trench. So the trench Tesoro 102 would get shortened to T102, the trench Civitate B 25 would be CB25, and so on.
        Each row should be the coordinates in the local master grid (NOT WGS84 or ESPG3003) The trench should start with the NW corner and move clockwise around the trench. so a standard rectangular trench would be NW NE SE SW.
        save that file and place it in the same folder as compile_forms.py



What is compile_forms.py?
    This is the main python file that solves the above problem. Read in the kobo form xlsx. This is a excel file (workbook) that has multiple internal sheets used for organization. Each sheet gets read in as a pandas dataframe. Every locus is processed, will all appropriate fields being found through the excel sheet. Each section is formatted correctly for the form. Then through the use of mailmerge, a new form is made, populated, and saved in the appropriate spot.


What is Mail Merge?
    Mail Merge is a system that allows for populating copies of documents on mass (https://en.wikipedia.org/wiki/Mail_merge)

What is Trench_Coordinates.csv?
    The only piece of information needed for the US form that is not stored in the kobo form are the coordinates for each trench. These coordinates need to be in the master grid for Poggio Civitate, and NOT in a known CRS like WGS84 or ESPG3003 (The CRS for our GIS)
    The format of the csv should be as follows:
        Each column header should be the shortened designation for the trench. So the trench Tesoro 102 would get shortened to T102, the trench Civitate B 25 would be CB25, and so on.
        Each row should be the coordinates in the local master grid (NOT WGS84 or ESPG3003) The trench should start with the NW corner and move clockwise around the trench. so a standard rectangular trench would be NW NE SE SW.
        save that file and place it in the same folder as compile_forms.py
    in practice the csv could look something like this (example comes from 2023):
    T90,T101,T103,CA93,CA94,CA95,T102,T104
    NW: 103/-38,NW: 169/-41.12,NW: 177/-64,NW: 27/12,NW1: 19/17; NW2:21/16,NW: 44/-16,NW: 190/-53,NW: 191/53
    NE: 110/-38,NE: 171/-41.90,NE: 182/-64,NE: 30/12,NE1: 21/17; NE2: 23/16,NE: 47/-16,NE: 191/-53,NE: 194/53
    SE: 110/-48,SE: 171/-44,SE: 182/-67,SE: 30/09,SE1: 21/12; SE2: 23/13,SE: 47/-20,SE: 191/-56,SE: 194/56
    SW: 103/-38,SW: 169/-44,SW1: 179/-67,SW: 27/09,SW1: 19/12; SW2: 21/13,SW: 44/-20,SW: 190/-56,SW: 191/56
    ,,SW2: 179/-66,,,,,
    ,,SW3: 177/-66,,,,,
        note the different coordinate naming conventions, it is okay to have multiple. N and E are positive directions, S and W are negative.

What is kobo_input_all.xlsx?
    This is all the data


The column headers for the Locus Summary sheet of kobo_input_all.xlsx
#current row headers:
            '''
            {
            'start': '2023-07-08T16:57:26.013+02:00', 
            'end': '2023-07-16T11:13:17.716+02:00', 
            'Data_creator': 'KRK',
            'Data_creator_txt': '',
            'Trench': 't90_2023',
           'Trench_ID': 't90_2023', 
            'Locus_ID': '41',
           'Season': '2023',
            'OC_Locus_ID': 'Locus 41', 
            'Trench_Supervisor': 'KRK',
           'Trench_Supervisor_txt': '',
            'Description': 'Overall Description',
           'Date_Opened': '2023-07-05',
            'Date_Closed': '2023-07-06', 
            'Locus_Type': 'deposit',
           'Locus_Type_Other': '',
           'Preliminary_Phasing': 'undetermined', 
            'Preliminary_Phasing/undetermined': '1',
           'Preliminary_Phasing/-900_-800': '0', 
            'Preliminary_Phasing/-800_-700': '0',
           'Preliminary_Phasing/-700_-600': '0', 
            'Preliminary_Phasing/-600_-500': '0',
           'Preliminary_Phasing/-500_-400': '0', 
            'Preliminary_Phasing/-400_-300': '0',
           'Preliminary_Phasing/-300_-200': '0', 
            'Preliminary_Phasing/-200_-100': '0',
           'Preliminary_Phasing/-100_0': '0', 
            'Preliminary_Phasing/Late': '0', 
            'Munsell_Color': '2.5Y 4/4 olive brown', 
            'Deposit_Compaction': 'medium',
           'Stratigraphic_Reliability': 'fair', 
            'Natural_or_Artificial': 'artificial', 
            'Recovery_Techniques': '', 
            'Recovery_Techniques/Hand_Picked': '',
           'Recovery_Techniques/Dry_Sieved': '', 
            'Recovery_Techniques/Wet_Sieved': '',
           'Recovery_Techniques/Flotation': '', 
            'Recovery_Techniques_Other': '', 
            'Mode_of_Formation': 'Anthropic', 
            'Inorganic_Components': '', 
            'Organic_Components': '', 
            'Definition_and_Position': "Inside the Trench", 
            'Criteria_for_Distinction': '', 
            'State_of_Conservation': 'Removed following documentation', 
            'Observations': '',
           'Interpretation': '',
            'Locus_Period': '', 
            'Locus_Period/Iron_Age': '',
           'Locus_Period/Early_Orientalizing': '', 
            'Locus_Period/Orientalizing': '',
           'Locus_Period/Archaic': '', 
            'Locus_Period/Classical': '',
           'Locus_Period/Hellenistic': '', 
            'Locus_Period/Roman': '',
           'Locus_Period/Post_Antique': '',
            'Locus_Period_Other': '', 
            'Length_m': '0.73',
           'Width_m': '0.38',
            'Depth_cm': '5',
           'Area_m_sq': '0.28', 
            'Authority': '',
           'Location': '',
            'Director': '',
           'Field_Director': '', 
            'Skip_Stratigraphy': 'AUTOMATIC',
           '_1_Cuts': '',
           '_2_Is_Cut_By': '',
           '_3_Fills': '',
           '_4_Is_Filled_By': '',
           '_5_Is_Bound_To': '',
           '_6_Same_As': '',
           '_7_Above': '',
           '_8_Below': '',
           '_9_Overlies': '',
           '_10_Underlies': '',
           '_11_Anterior_To': '',
           '_12_Posterior_To': '', 
            'Strat_note': '',
           'Note_the_Trench_ID_f_us_in_another_trench': '',
           'For_more_information_types': '',
            'Image': 'Photo-2023-T90-Locus41-005-15_3_42.jpg',
           'Image_URL': 'https://kcat.opencontext.org/media/original?media_file=eric%2Fattachments%2F9cfd658f785645b495616b8a3b8e3584%2F36f60f55-72fc-4906-acf1-159a314eabfd%2FPhoto-2023-T90-Locus41-005-15_3_42.jpg',
           'Image_Note': 'Locus 41 opening photo, from E looking W',
            '_1_Cuts_The_creatio_tified_in_this_field': '',
           '_2_Overlies_This_lo_tified_in_this_field': '',
           '_3_Above_This_locus_in_physical_contact': '',
           '_4_Below_This_locus_in_physical_contact': '',
           '_5_Abuts_This_locu_tified_in_this_field': '',
           '_6_Contemporary_with_nce_at_the_same_time': '',
           '_7_Same_as_The_locu_scribed_on_this_form': '',
           'For_more_information_ypes_of_relationship': '',
           '_id': '135',
           '_uuid': 'f32e2872-d711-4163-a2d2-b7d41e87319e', 
            '_submission_time': '2023-07-08T15:17:17',
           '_validation_status': '',
           '_notes': '',
           '_status': 'submitted_via_web',
           '_submitted_by': '',
           '__version__': 'v3VkLLYpLakofFjAt9r5iV',
           '_tags': '',
            '_index': '1'
            }
            '''

Current mailmerge spots in the template docx file
    {'Interpretations', 'Natural', 'Flotation', 'Is_Bound_To', 'Artificial', 'Elevation_or_Depth', 'Following_Loci', 'Leans_On', 'Locus_Number_2', 'Inorganic_List', 'Equal_To', 'Quantity_Report', 'Sieving', 'Dimensions', 'Criteria_for_Distinction', 'Is_Filled_By', 'Peroid_or_Phase', 'Covers', 'Reliability', 'Samples', 'Locus_Number', 'Area', 'Color', 'Trench_Coordinates', 'Definition_and_Position', 'Previous_Loci', 'Photo_Names', 'Is_Cut_By', 'Catalogue_Numbers', 'Organic_List', 'Year', 'Mode_of_Formation', 'Plan_Number', 'Fills', 'Cuts', 'Description', 'Datable_Elements', 'Covered_By', 'State_of_Conservation', 'Observations', 'Is_Above', 'Consistency', 'Trench_Number', 'Dating'}

Current sheets in the full .xlsx book, I'm hoping these names don't change with the seasons, but they very well could
#current sheet names:
        #'Locus Summary Entry 2023',   -- main sheet
        #'group_trench_book',          -- Trench book page references (NOT NEEDED)
        #'group_strat_other',          -- Relates Loci to each other (EQUAL TO?)
        #'group_elevations',           -- Elevations + Coordinates for individual Locus
        #'group_files',                -- Unknown, seems to be uploaded photos (NOT NEEDED)
        #'begin_repeat_Et0mtvMid',     -- Associated Plans
        #'begin_repeat_NawXUcMgh',     -- Associated Photos
        #'begin_repeat_pyQUfbbMJ',     -- Special Finds
        #'begin_repeat_It8bliv1T',     -- Tabulated counts
        #'begin_repeat_UlorltmZm',     -- Datable Elements 
        #'begin_repeat_v62FStDQL'      -- Associated Samples (Contains_Samples)
        #'group_strat_Above',          -- 
        #'group_strat_Overlies',       --
        #'group_strat_Underlies',      --
        #'group_strat_Is_Cut_By'       --
        #'group_strat_Cuts'            --
        #'group_strat_Same_As'         --
        #'group_strat_Is_Bound_To'     --
        #'group_strat_Fills'           --
        #'group_strat_Is_Filled_By'    --
        #'group_strat_Below',          --  
        #'group_strat_Anterior_To',    -- What comes after numerically (documented locus 2 so 3,4 come after)
        #'group_strat_Posterior_To'    -- What comes before numerically (documented locus 2 so 1 came before)