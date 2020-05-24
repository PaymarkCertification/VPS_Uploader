__author__ = ["Michaely"]
import pandas as pd
import os
import Config as c
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
col = ['Summary', 'Cert Defect Class', 'Labels', 'Status', 'Components', 'Description', 'Key', 'Issue Type']
server = c.webDriver["address"]
options = {
 'server': server
}

try:
    jira = JIRA(options, basic_auth=(c.jira["user"], c.jira["apiKey"]))
    log.info("Authenticating OAuth")
except JIRAError as e:
    if e.status_code == 401:
        log.info("Unable to authenticate")
    else:
        log.info("unknown error")
    log.error('Issue with auth using OAuth params')


def retrieveKeys():
    header = "Summary, Key, Issue Type "
    # return issues from JIRA
    issues_in_proj = jira.search_issues(jql_str=c.jira["JQL"],
                                        fields='summary, key',
                                        maxResults=999,
                                        expand=True,
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
    issues_in_proj = jira.search_issues(jql_str=c.jira["JQL"],
                                        maxResults=999,
                                        expand=True,
                                        json_result=True)
    print(json.dumps(issues_in_proj, indent=4))


def newIssue():
    issueList = []
    noKey = []
    count = 0
    with open('upload1.json') as f:
        data = json.load(f)
    log.info('Converting JSON to Python Object')
    for x in range(len(data)):
            x-1
            if data[x]['Key'] is None:
                noKey.append(data[x])
                for i in range(len(noKey)):
                    issue_dict = {
                        'project': {col[6]: 'TAC'},
                        'summary': data[x][col[0]],
                        'description': data[x][col[5]],
                        'issuetype': {'name': 'Defect'},
                        'customfield_12604': {'value': data[x][col[1]]},
                        'components': [{'name': data[x][col[4]]}],
                        'labels': [data[x][col[2]]]
                    }
                    count += 1
                    issueList.append(issue_dict)
                else:
                    log.error('Key index out of range')
    jira.create_issues(field_list=issueList)
    log.info("Creating {} Issue(s)".format(count))
    issues_in_proj = jira.search_issues(jql_str=c.jira["JQL"], fields='summary, key', maxResults=999, expand=True,
                                        json_result=False)
    a = 0
    try:
        for issues in issues_in_proj:
            a += 1
        log.info("Returned {} Issues".format(a))
    except JIRAError as e:
        log.error("JQL returned no results", e)


def updateIssue():
    wKey = []
    count = 0
    with open('upload1.json') as f:
        data = json.load(f)
        for i in range(len(data)):
            if data[i]['Key'] is not None:
                wKey.append(data[i])
        for x in range(len(wKey)):
                    x-1
                    count += 1
                    jira.issue(wKey[x]['Key']).update(summary=wKey[x]['Summary'], description=wKey[x]['Description'])
    log.info("Updating {} Issues".format(count))



def excelParser():
    # parse excel sheet at spec. column index [Fault ID, Classification, Project Ref, Status, Phase, Issue]
    try:
        df1 = pd.read_excel(c.VPS["file"],
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


# clean up func
def exitOpt():
    input("Press 'ENTER' to quit")
    os.remove("upload1.json")
    os.remove(req)
    quit()


if __name__ == "__main__":
    print("VPS Uploader\nVersion: 0.2 \nRelease Notes:\n--Issue update function added\n"
          "--Users can confirm JQL & VPS prior to runtime\n--Added OAuth Exception\n")
    print("JQL string used: {}".format(c.jira["JQL"]))
    print("VPS: {}".format(c.VPS["file"]))
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
            # newIssue()
            # updateIssue()
            print('issue')
            exitOpt()
        elif cmd.lower() == 'n':
            exitOpt()
    if choice.lower() == 'n':
        log.info("Ending Script")
        exitOpt()
    if choice.lower() == 'exit':
        log.info("exiting")
        quit()
