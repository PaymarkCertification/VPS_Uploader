__author__ = ["Michaely"]

import sys
import os
sys.path.append(os.getcwd()+'/'+'library'+'/')
import re
from library import pandas as pd
from os import listdir
import Config as C
from library.jira import JIRA, JIRAError
import json
import logging



'''logging variables'''
log = logging.getLogger()
console = logging.StreamHandler()
format_str = '%(asctime)s\t%(levelname)s -- %(processName)s %(filename)s:%(lineno)s -- %(message)s'
console.setFormatter(logging.Formatter(format_str))
log.addHandler(console)
log.setLevel(logging.INFO)

'''random variables'''
req = 'Jira_results.txt'
col = ['Summary', 'Cert Defect Class', 'Labels', 'Status', 'Components', 'Description', 'key', 'Issue Type']
parentFolder = os.getcwd()
server = C.webDriver["address"]
options = {
 'server': server
}


def setupExcelFile():
    excelFilepath = [f for f in listdir(parentFolder) if f.endswith('.xlsm')]
    if not excelFilepath:
        log.error("No excel found in {}".format(parentFolder))
        exitOpt()
    elif len(excelFilepath) > 1:
        log.error("Error: More than one excel found. Can only parse one file.")
        exitOpt()
    return excelFilepath[0]


# clean up func
def exitOpt():
    txt = C.jira['user']
    x = re.split("[$@]", txt)
    y = re.sub("[$.]", ' ', x[0])
    z = re.split("\s", y)
    print('Thanks for using the script', z[0], '!')
    input("Press 'ENTER' to quit")
    log.info('--Ending Script')
    if os.path.exists("upload1.json"):
        os.remove("upload1.json")
    if os.path.exists(req):
            os.remove(req)
    else:
        pass
    quit()


try:
    jira = JIRA(options, basic_auth=(C.jira["user"], C.jira["apiKey"]))
    log.info("Authenticating OAuth")
except JIRAError as e:
    if e.status_code == 401:
        log.info("Unable to authenticate. Check OAuth credentials")
        exitOpt()
    elif e.status_code == 403:
        log.info("User is not an administrator")
        exitOpt()
    else:
        log.info("Error: ", e.status_code)
        exitOpt()
    log.error('Issue with auth using OAuth params')


def retrieveKeys():
    header = "Summary, Key, Issue Type "
    # return issues from JIRA
    issues_in_proj = jira.search_issues(jql_str=C.jira["JQL"], fields='summary, key', maxResults=999, expand=True,
                                        json_result=False)
    log.info("Querying JQL")
    x = 0
    output = open(req, 'w')
    output.write(header+"\n")
    for i in issues_in_proj:
        x += 1
        print(i.fields.summary, ',', i.key, ',', 'Defect', file=output)
    log.info("Returned {} Issues".format(x))
    log.info("Writing Jira Issues to CSV")


def allIssueFields():
    issues_in_proj = jira.search_issues(jql_str=C.jira["JQL"], maxResults=999, expand=True, json_result=True)
    print(json.dumps(issues_in_proj, indent=4))


def newIssue():
    issueList = []
    noKey = []
    with open('upload1.json') as f:
        data = json.load(f)
    log.info('Converting JSON to Python Object')
    for x in range(len(data)):
            if data[x][col[6]] is None:
                noKey.append(data[x])
    for i in range(len(noKey)):
            issue_dict = {
                            'project': {col[6]: 'TAC'},
                            'summary': noKey[i][col[0]],
                            'description': noKey[i][col[5]],
                            'issuetype': {'name': 'Defect'},
                            'customfield_12604': {'value': noKey[i][col[1]]},
                            'components': [{'name': noKey[i][col[4]]}],
                            'labels': [noKey[i][col[2]]]
                            }
            issueList.append(issue_dict)
    jira.create_issues(field_list=issueList)
    log.info("Creating {} Issue(s)".format(len(issueList)))

    a = 0
    try:
        log.info('Querying JQL post upload')
        issues_in_proj = jira.search_issues(jql_str=C.jira["JQL"], fields='summary, key', maxResults=999, expand=True,
                                            json_result=False)
        for issues in issues_in_proj:
            a += 1
        log.info("JQL returned {} Issue(s)".format(a))
    except JIRAError as error:
        log.error("JQL returned no results", error)


def updateIssue():
    wKey = []
    count = 0
    log.info('Updating')
    with open('upload1.json') as f:
        data = json.load(f)
        for i in range(len(data)):
            if data[i][col[6]] is not None:  # if key exists update
                wKey.append(data[i])
        for x in range(len(wKey)):
            count += 1
            jira.issue(wKey[x]['key']).update(summary=wKey[x]['Summary'], description=wKey[x]['Description'])
    log.info("Updated {} Issues".format(count))



def excelParser():
    # parse excel sheet at spec. column index [Fault ID, Classification, Project Ref, Status, Phase, Issue]
    df1 = pd.read_excel(setupExcelFile(), sheet_name='Issue List', usecols='C,D,F,G,H,I')
    log.info('Parsing excel')
    # 'Summary', 'Cert Defect Class', 'Labels', 'Status', 'Components', 'Description', 'Key', 'Issue Type'
    df1.columns = [col[0], col[1], col[2], col[3], col[4], col[5]]
    df1.dropna(subset=[col[0]], inplace=True)
    df1.sort_values(by=col[0], ascending=True)
    df2 = pd.read_csv(req)
    log.info("Parsing CSV")
    df2.columns = [col[0], col[6], col[7]]
    df2[col[0]] = df2[col[0]].str.strip()
    df2[col[6]] = df2[col[6]].str.strip()
    df3 = df1.merge(df2, on=[col[0]], how='left')
    log.info("Merging on column {}".format(col[0]))
    # df3['Key'].replace([None], value=np.nan, inplace=True)
    # df3['Key'] = df3['Key'].fillna("nokey", inplace=True)
    df3.to_json('upload1.json', orient='records')
    log.info("Creating JSON")
    print('\n#########################################################################################################'
          '#########')
    print(df3)
    print('##########################################################################################################'
          '########')
    print("[Note]: Issues with Keys will be updated, those without are created as new issues.\n")


if __name__ == "__main__":
    with open(parentFolder+'/'+'library'+'/'+"Release Notes.txt", "r") as j:
        print(j.read())
    print("\n<<Parameters:>>")
    print("JQL string: {}".format(C.jira["JQL"]))
    print("VPS: {}".format(setupExcelFile()))
    while True:
        choice = input("\nIs the information correct? (Y/N): ")
        if choice.lower() not in ('y', 'n', 'exit'):
            print("unrecognised key entered. Y Or N?")
        else:
            break
    if choice.lower() == 'y':
        retrieveKeys()
        excelParser()
        while True:
            cmd = input("Press 'Y' to continue... ")
            if cmd.lower() not in ('y', 'n', 'exit'):
                print("Unrecognised command\nCMD? ")
            else:
                break
        if cmd.lower() == 'y':
            print("Processing...")
            newIssue()
            updateIssue()
            print("Completed")
            exitOpt()

        if cmd.lower() == 'n':
            exitOpt()

    if choice.lower() == 'n':
        exitOpt()

    if choice.lower() == 'exit':
        exitOpt()
