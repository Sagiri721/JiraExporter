from jira import JIRA
import json
import xlsxwriter;
import sys

if __name__ == '__main__':    

    print("\n### JIRA backlog saver started! ###");
    print("READ HELP.MD!!!!");
    print("Press any key to continue...");
    input();

    print("\nPlease provide your Jira credentials and project key to save the backlog issues.");
    username: str = input("Username: ");
    api_token: str = input("API Token: ");
    project_key: str = input("Project Key: ");
    server: str = input("Jira Server: ");

    # Output information
    my_issues = {};

    # Connect to Jira
    print("Attempting to connect to Jira...");
    jira = JIRA(server=server, basic_auth=(username, api_token));
        
    print("Searching for issues in the backlog...");
    backlog_issues = jira.search_issues(f'project={project_key} AND Sprint is EMPTY');

    print("Parsing backlog issues...")
    for issue in backlog_issues:
        
        my_issues[issue.key] = {
            'summary': issue.fields.summary,
            'description': issue.fields.description,
            'type': issue.fields.issuetype.name,
            'priority': issue.fields.priority.name,
            'status': issue.fields.status.name,
            'assignee': issue.fields.assignee.displayName if issue.fields.assignee is not None else 'Unassigned',
            'reporter': issue.fields.reporter.displayName,
            'created': issue.fields.created,
        };

    # Write json file
    print("Saving raw issues to JSON...");
    with open('backlog_issues.json', 'w+') as file:
        json.dump(my_issues, file);
    
    print("Backlog issues saved to 'backlog_issues.json' file.");

    # Write xlsx file
    print("Saving issues to XLSX...");
    workbook = xlsxwriter.Workbook('backlog_issues.xlsx');
    worksheet = workbook.add_worksheet()

    # Writing headers
    worksheet.write('A1', 'Issue Key');
    worksheet.write('B1', 'Summary');
    worksheet.write('C1', 'Description');
    worksheet.write('D1', 'Type');
    worksheet.write('E1', 'Priority');
    worksheet.write('F1', 'Status');
    worksheet.write('G1', 'Assignee');
    worksheet.write('H1', 'Reporter');
    worksheet.write('I1', 'Created');
    
    # Writing issues
    row = 2;
    for key, issue in my_issues.items():
        worksheet.write(f'A{row}', key);
        worksheet.write(f'B{row}', issue['summary']);
        worksheet.write(f'C{row}', issue['description']);
        worksheet.write(f'D{row}', issue['type']);
        worksheet.write(f'E{row}', issue['priority']);
        worksheet.write(f'F{row}', issue['status']);
        worksheet.write(f'G{row}', issue['assignee']);
        worksheet.write(f'H{row}', issue['reporter']);
        worksheet.write(f'I{row}', issue['created']);
        row += 1;

    workbook.close();

    print("Excel file created! Check 'backlog_issues.xlsx' for the backlog issues.");
    print("\n### JIRA backlog saver finished! ###");
