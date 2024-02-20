import os
import pandas as pd

os.system('pip install -r requirements.txt')

output = pd.read_excel('data/output_template.xlsx')
inputDf = pd.read_excel('data/enrollees_list.xlsx')


def to_camel_case(name):
    words = name.split()
    if not words:
        return ""
    ccw = [word.capitalize() + " " for word in words]
    return ''.join(ccw)


inputDf["Register Number"] = inputDf.loc[inputDf["Register Number"].notna(), "Register Number"].apply(lambda x: x[:2])
inputDf["Student Name"] = (inputDf["Student Name"].fillna("") + " " + inputDf["Register Number"].fillna(""))\
    .apply(to_camel_case)
inputDf = inputDf[(inputDf["Student Name"] != "") & (inputDf["Student Name"] != " ")]

output["Name"] = inputDf["Student Name"]
output["Given Name"] = inputDf["Student Name"]
output["Group Membership"] = "* myContacts"
output["E-mail 1 - Type"] = "Work"
output["E-mail 1 - Value"] = inputDf["EmailId"]
output["Phone 1 - Type"] = "Mobile"
output["Phone 1 - Value"] = inputDf["Mobile No"]

output.to_csv('data/contacts.csv', index=False)
