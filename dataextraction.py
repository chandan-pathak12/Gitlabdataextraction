import gitlab
import pandas as pd
from datetime import timedelta

# GitLab setup
GITLAB_URL = 'https://git.impressicocrm.com'
PRIVATE_TOKEN = 'glpat-XXXXXXXXXXXXXXXXXG4Kz'
gl = gitlab.Gitlab(GITLAB_URL, private_token=PRIVATE_TOKEN)

GROUP_NAME = 'devOps'
PROJECT_NAME = 'devOps'

START_DATE = pd.to_datetime('2024-12-01')
END_DATE = pd.to_datetime('2024-12-31')

# Define active and tracked labels
ACTIVE_LABELS = [
    'Ready for Progress', 'In Progress', 'Documentation', 
    'Ready to Test', 'Validation', 'Mentor Review', 
    'Lead Review', 'Review', 'Ready to Merge', 'Done',
    'Priority-1', 'Priority-2', 'Priority-3'
]
TRACKED_LABELS = ['In Progress', 'Lead Review', 'Review', 'Mentor Review']

# Function to fetch GitLab project
def get_project():
    group = gl.groups.get(GROUP_NAME)
    return gl.projects.get(f'{GROUP_NAME}/{PROJECT_NAME}')

# Function to filter issues by date
def filter_issues_by_date(issues, start_date, end_date, project):
    return [
        project.issues.get(issue.iid) 
        for issue in issues 
        if start_date <= pd.to_datetime(issue.created_at).tz_localize(None) <= end_date
    ]

# Function to process issues and prepare data
def process_issues(filtered_issues):
    data = []
    creator_count = {}

    for issue in filtered_issues:
        created_at = pd.to_datetime(issue.created_at).tz_localize(None)
        closed_at = pd.to_datetime(issue.closed_at).tz_localize(None) if issue.closed_at else None
        assignee_name = issue.assignee['name'] if issue.assignee else 'Unassigned'
        creator_name = issue.author['name']
        lifespan = (closed_at - created_at).days if closed_at else None
        time_stats = issue.time_stats()
        time_spent = time_stats['total_time_spent'] / 3600 if time_stats else 0
        time_estimate = time_stats['time_estimate'] / 3600 if time_stats else 0

        # Increment the issue count for the creator
        creator_count[creator_name] = creator_count.get(creator_name, 0) + 1

        time_spent_in_labels = {label: timedelta(0) for label in TRACKED_LABELS}
        events = issue.resourcelabelevents.list(all=True)
        label_start_times = {}

        for event in events:
            event_time = pd.to_datetime(event.created_at).tz_localize(None)
            label_name = event.label['name'] if event.label else None
            if label_name in TRACKED_LABELS:
                if event.action == 'add':
                    label_start_times[label_name] = event_time
                elif event.action == 'remove' and label_name in label_start_times:
                    time_spent_in_labels[label_name] += event_time - label_start_times.pop(label_name)

        for label, duration in time_spent_in_labels.items():
            time_spent_in_labels[label] = duration.days  # Convert to days

        issue_data = {
            'Issue ID': issue.iid,
            'Title': issue.title,
            'Assignee': assignee_name,
            'Creator': creator_name,
            'Lifespan (days)': lifespan,
            'Time Spent (hours)': time_spent,
            'Time Estimate (hours)': time_estimate,
            'Start Date': created_at.strftime('%Y-%m-%d'),
            'End Date': closed_at.strftime('%Y-%m-%d') if closed_at else None
        }

        for label in TRACKED_LABELS:
            issue_data[f'Time Spent in {label} (days)'] = time_spent_in_labels[label]

        for label in ACTIVE_LABELS:
            issue_data[label] = label in issue.labels

        issue_data['Closed in Period'] = closed_at and START_DATE <= closed_at <= END_DATE
        data.append(issue_data)

    return pd.DataFrame(data), creator_count

# Main function
def main():
    project = get_project()
    issues = project.issues.list(all=True)
    filtered_issues = filter_issues_by_date(issues, START_DATE, END_DATE, project)
    df, creator_count = process_issues(filtered_issues)

    priority_1_df = df[df['Priority-1']]
    priority_2_df = df[df['Priority-2']]
    priority_3_df = df[df['Priority-3']]

    # Save the issue details and creator counts
    with pd.ExcelWriter('filtered_issues_with_time_spent_in_labels.xlsx') as writer:
        df.to_excel(writer, sheet_name='All Issues', index=False)
        priority_1_df.to_excel(writer, sheet_name='Priority-1', index=False)
        priority_2_df.to_excel(writer, sheet_name='Priority-2', index=False)
        priority_3_df.to_excel(writer, sheet_name='Priority-3', index=False)

    creator_count_df = pd.DataFrame(list(creator_count.items()), columns=['Creator', 'Issue Count'])
    creator_count_df.to_excel('issue_count_by_creator.xlsx', index=False)

    print("Done! Issue details and creator counts saved.")

# Run the script
if __name__ == '__main__':
    main()
