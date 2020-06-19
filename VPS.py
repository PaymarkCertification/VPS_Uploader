__author__ = ["Michaely"]
from lib import *
# import pandas as pd
import os
import Config as C
from jira import JIRA, JIRAError
import json
import logging

log = logging.getLogger()
console = logging.StreamHandler()
format_str = '%(asctime)s\t%(levelname)s -- %(processName)s %(filename)s:%(lineno)s -- %(message)s'
console.setFormatter(logging.Formatter(format_str))
log.addHandler(console)
log.setLevel(logging.INFO)
req = 'Jira_results.txt'
col = ['Summary', 'Cert Defect Class', 'Labels', 'Status', 'Components', 'Description', 'key', 'Issue Type']
server = C.webDriver["address"]
options = {
 'server': server
}

try:
    jira = JIRA(options, basic_auth=(C.jira["user"], C.jira["apiKey"]))
    log.info("Authenticating OAuth")
except JIRAError as e:
    if e.status_code == 401:
        log.info("Unable to authenticate. Check OAuth credentials")
    elif e.status_code == 403:
        log.info("User is not an administrator")
    else:
        log.info("unknown error")
    log.error('Issue with auth using OAuth params')


def retrieveKeys():
    header = "Summary, Key, Issue Type "
    # return issues from JIRA
    issues_in_proj = jira.search_issues(jql_str=C.jira["JQL"], fields='summary, key', maxResults=999, expand=True,
                                        json_result=False)
    log.info("Returning JQL query")
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
    log.info('Converting JSON')
    log.info('Processing Upload...')
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
        issues_in_proj = jira.search_issues(jql_str=C.jira["JQL"], fields='summary, key', maxResults=999, expand=True,
                                            json_result=False)
        for issues in issues_in_proj:
            a += 1
        log.info("Returned {} Issues".format(a))
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
                    jira.issue(wKey[x]['Key']).update(summary=wKey[x]['Summary'], description=wKey[x]['Description'])
    log.info("Updated {} Issues".format(count))



def excelParser():
    # parse excel sheet at spec. column index [Fault ID, Classification, Project Ref, Status, Phase, Issue]
    try:
        df1 = pd.read_excel(C.VPS["file"],
                            sheet_name='Issue List',
                            usecols='C,D,F,G,H,I')
        log.info('Parsing excel')
    except FileNotFoundError:
        log.error("Wrong file or file path")
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
    print('##########################################################')
    print(df3)
    print('##########################################################')
    print("[Note]: Issues with Keys will be updated, those without are created as new issues.")


# clean up func
def exitOpt():
    input("Press 'ENTER' to quit")
    os.remove("upload1.json")
    os.remove(req)
    quit()


if __name__ == "__main__":
    print("#######")
    print("VPS Uploader\nVersion: 0.3 \nRelease Notes:\n--Issue update function added\n"
          "--Users can confirm JQL & VPS prior to runtime\n--Added OAuth Exception handling\n--fixed createNew"
          "\n--Updated lib import")
    print("#######\n")
    print("JQL string: {}".format(C.jira["JQL"]))
    print("VPS: {}".format(C.VPS["file"]))
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
        elif cmd.lower() == 'n':
            exitOpt()
    if choice.lower() == 'n':
        log.info("Ending Script")
        exit()
    if choice.lower() == 'exit':
        log.info("exiting")
        quit()
