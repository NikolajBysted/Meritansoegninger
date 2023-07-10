import pandas as pd

# Load your Excel document
df = pd.read_excel('input.xlsx')

# Column names
col_names = ["Dato for ansøgning", "Studienummer", "Dato for afgørelse", "Semester", "Status", "Land", "Universitet", "Kursusnr./Fagkode", "Kursusnavn", "credits", "ects", "Overføres som (fagområde)", "Evt. obligatorisk kursus", "Niveau", "Uddannelse", "Årsværk", "Godkendt/Afvist", "Begrundelse for afvisning", "FV", "Sagsbehandler", "Type af ophold", "Studerendes bemærkning", ]
result_rows = [col_names]

for idx, row in df.iterrows():
    
    number_of_applications = 0
    for i in range(0, 10):
        cell = row[i + 10]
        if not (pd.isnull(cell)):
            number_of_applications += 1

    # Get info about the person
    application_date = row[6]
    study_number = row[171] 
    date_of_decision = ""
    semester = str(row[54]) + " " + str(int(row[110]))
    status = ""
    country = "" #ifølge meritskabelon
    education_org = row[47]
    education = row[50]
    accepted_or_declined = ""
    reason = ""
    fv = ""
    sagsbehanlder = ""
    type_of_stay = row[51]
    comment= row[172]

    for i in range(0, number_of_applications):

        # Get info about application
        course_number = row[24 + i]
        course_name = row[131 + i]
        choice = row[10 + i]

        credits = ""
        ects = ""
        
        if (choice == "vaelg_credits"):
            credits = row[111 + i]
        else:
            ects = row[111 + i]

        #if (choice == "valg_ects"):
        #    ects_or_credits = ects
        #else:
        #    ects_or_credits = credits
        
        transferred_as = row[157 + i]
        mand_course_name = row[78 + i]
        level = row[146 + i]
        fte = row[121 + i]

        result_row = [application_date, 
                      study_number, 
                      date_of_decision, 
                      semester, 
                      status, 
                      country, 
                      education_org, 
                      course_number,
                      course_name,
                      credits,
                      ects,
                      transferred_as,
                      mand_course_name,
                      level,
                      education,
                      fte,
                      accepted_or_declined,
                      reason,
                      fv,
                      sagsbehanlder,
                      type_of_stay,
                      comment]

        result_rows.append(result_row)




    
    #print(number_of_applications)

#print(result_rows)

df_out = pd.DataFrame(result_rows)
df_out.reset_index(drop=True).to_excel('output.xlsx', index=False)