import requests
import pandas as pd
import os
import sys
import json

TOKEN = os.environ.get("GITHUB_TOKEN")
PROJECT_ID = os.environ.get("PROJECT_ID")

if not TOKEN or not PROJECT_ID:
    print("Missing GITHUB_TOKEN or PROJECT_ID")
    sys.exit(1)

url = "https://api.github.com/graphql"

headers = {
    "Authorization": f"Bearer {TOKEN}"
}

query = """
query ($projectId: ID!) {
  node(id: $projectId) {
    ... on ProjectV2 {
      items(first: 100) {
        nodes {
          content {
            ... on Issue {
              number
              title
              state
            }
          }
          fieldValues(first: 20) {
            nodes {
              ... on ProjectV2ItemFieldNumberValue {
                number
                field {
                  ... on ProjectV2FieldCommon {
                    name
                  }
                }
              }
              ... on ProjectV2ItemFieldTextValue {
                text
                field {
                  ... on ProjectV2FieldCommon {
                    name
                  }
                }
              }
            }
          }
        }
      }
    }
  }
}
"""

response = requests.post(
    url,
    json={"query": query, "variables": {"projectId": PROJECT_ID}},
    headers=headers
)

data = response.json()

# Debug errors clearly
if "errors" in data:
    print("GraphQL Errors:")
    print(json.dumps(data["errors"], indent=2))
    sys.exit(1)

items = data.get("data", {}).get("node", {}).get("items", {}).get("nodes", [])

rows = []

for item in items:
    issue = item.get("content")
    if not issue:
        continue

    row = {
        "ID": issue.get("number"),
        "Title": issue.get("title"),
        "State": issue.get("state"),
        "Original Estimate Days": None,
        "Time Spent Days": None,
        "Remaining Estimate Days": None
    }

    for field in item.get("fieldValues", {}).get("nodes", []):
        field_info = field.get("field")
        if not field_info:
            continue

        field_name = field_info.get("name")
        value = field.get("number") or field.get("text")

        if field_name == "Original Estimate Days":
            row["Original Estimate Days"] = value
        elif field_name == "Time Spent Days":
            row["Time Spent Days"] = value
        elif field_name == "Remaining Estimate Days":
            row["Remaining Estimate Days"] = value

    rows.append(row)

if not rows:
    print("No data found. Check PROJECT_ID or field names.")
    sys.exit(1)

df = pd.DataFrame(rows)

output_file = "github_project_report.xlsx"
df.to_excel(output_file, index=False)

print(f"Exported {len(df)} rows to {output_file}")
