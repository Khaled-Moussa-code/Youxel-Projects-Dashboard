"""
Sprint Analysis Automation - Notion Exporter Module

This module handles exporting sprint data to Notion via API.
"""

import json
import requests
import pandas as pd
from typing import Dict, List
from datetime import datetime


class NotionExporter:
    """Export sprint data to Notion databases"""
    
    def __init__(self, api_key: str, database_ids: Dict = None):
        """
        Initialize Notion exporter
        
        Args:
            api_key: Notion integration API key
            database_ids: Dict with database IDs for different data types
        """
        self.api_key = api_key
        self.headers = {
            "Authorization": f"Bearer {api_key}",
            "Content-Type": "application/json",
            "Notion-Version": "2022-06-28"
        }
        self.database_ids = database_ids or {}
    
    def create_sprint_summary_page(self,
                                   sprint_metadata: Dict,
                                   cmmi_measures: Dict) -> str:
        """
        Create a Sprint Summary page in Notion
        
        Returns:
            Notion page ID
        """
        if 'sprint_summary' not in self.database_ids:
            raise ValueError("Sprint Summary database ID not configured")
        
        payload = {
            "parent": {"database_id": self.database_ids['sprint_summary']},
            "properties": {
                "Sprint Name": {
                    "title": [
                        {
                            "text": {
                                "content": sprint_metadata['sprint_name']
                            }
                        }
                    ]
                },
                "Start Date": {
                    "date": {
                        "start": sprint_metadata['date_from'].isoformat()
                    }
                },
                "End Date": {
                    "date": {
                        "start": sprint_metadata['date_to'].isoformat()
                    }
                },
                "Completion Rate": {
                    "number": cmmi_measures['completion_rate']
                },
                "Utilization Rate": {
                    "number": cmmi_measures['utilization_rate']
                },
                "CMMI Productivity": {
                    "number": cmmi_measures['cmmi_productivity']
                },
                "Status": {
                    "select": {
                        "name": "Completed"
                    }
                }
            }
        }
        
        response = requests.post(
            "https://api.notion.com/v1/pages",
            headers=self.headers,
            data=json.dumps(payload)
        )
        
        if response.status_code == 200:
            return response.json()['id']
        else:
            raise Exception(f"Failed to create sprint page: {response.text}")
    
    def create_staff_performance_pages(self,
                                       staff_metrics: pd.DataFrame,
                                       sprint_id: str):
        """
        Create Staff Performance pages for each team member
        """
        if 'staff_performance' not in self.database_ids:
            raise ValueError("Staff Performance database ID not configured")
        
        created_pages = []
        
        for _, staff in staff_metrics.iterrows():
            payload = {
                "parent": {"database_id": self.database_ids['staff_performance']},
                "properties": {
                    "Staff Name": {
                        "title": [
                            {
                                "text": {
                                    "content": staff['staff_name']
                                }
                            }
                        ]
                    },
                    "Sprint": {
                        "relation": [
                            {
                                "id": sprint_id
                            }
                        ]
                    },
                    "Team": {
                        "select": {
                            "name": staff['team']
                        }
                    },
                    "KPI": {
                        "number": round(staff['kpi'], 4)
                    },
                    "Performance Rate": {
                        "number": round(staff['performance_rate'], 4)
                    },
                    "Utilization": {
                        "number": round(staff['utilization'], 4)
                    },
                    "Done Tasks %": {
                        "number": round(staff['done_tasks_pct'], 4)
                    },
                    "MidSprint Addition %": {
                        "number": round(staff['midsprint_addition_pct'], 4)
                    },
                    "Ad-hoc %": {
                        "number": round(staff['adhoc_pct'], 4)
                    }
                }
            }
            
            response = requests.post(
                "https://api.notion.com/v1/pages",
                headers=self.headers,
                data=json.dumps(payload)
            )
            
            if response.status_code == 200:
                created_pages.append(response.json()['id'])
            else:
                print(f"Warning: Failed to create page for {staff['staff_name']}: {response.text}")
        
        return created_pages
    
    def create_team_performance_pages(self,
                                      team_metrics: pd.DataFrame,
                                      sprint_id: str):
        """
        Create Team Performance pages
        """
        if 'team_performance' not in self.database_ids:
            raise ValueError("Team Performance database ID not configured")
        
        created_pages = []
        
        for _, team in team_metrics.iterrows():
            payload = {
                "parent": {"database_id": self.database_ids['team_performance']},
                "properties": {
                    "Team Name": {
                        "title": [
                            {
                                "text": {
                                    "content": team['team']
                                }
                            }
                        ]
                    },
                    "Sprint": {
                        "relation": [
                            {
                                "id": sprint_id
                            }
                        ]
                    },
                    "KPI": {
                        "number": round(team['kpi'], 4)
                    },
                    "Performance Rate": {
                        "number": round(team['performance_rate'], 4)
                    },
                    "Utilization": {
                        "number": round(team['utilization'], 4)
                    },
                    "Done Tasks %": {
                        "number": round(team['done_tasks_pct'], 4)
                    }
                }
            }
            
            response = requests.post(
                "https://api.notion.com/v1/pages",
                headers=self.headers,
                data=json.dumps(payload)
            )
            
            if response.status_code == 200:
                created_pages.append(response.json()['id'])
            else:
                print(f"Warning: Failed to create page for team {team['team']}: {response.text}")
        
        return created_pages
    
    def create_cmmi_history_page(self,
                                 cmmi_measures: Dict,
                                 sprint_id: str):
        """
        Create CMMI History page
        """
        if 'cmmi_history' not in self.database_ids:
            raise ValueError("CMMI History database ID not configured")
        
        payload = {
            "parent": {"database_id": self.database_ids['cmmi_history']},
            "properties": {
                "Sprint": {
                    "relation": [
                        {
                            "id": sprint_id
                        }
                    ]
                },
                "Completion Rate": {
                    "number": round(cmmi_measures['completion_rate'], 4)
                },
                "Effort Estimation Accuracy": {
                    "number": round(cmmi_measures['effort_estimation_accuracy'], 4)
                },
                "Bug Fixing %": {
                    "number": round(cmmi_measures['bug_fixing_effort_pct'], 4)
                },
                "Utilization Rate": {
                    "number": round(cmmi_measures['utilization_rate'], 4)
                },
                "CMMI Productivity": {
                    "number": round(cmmi_measures['cmmi_productivity'], 4)
                }
            }
        }
        
        response = requests.post(
            "https://api.notion.com/v1/pages",
            headers=self.headers,
            data=json.dumps(payload)
        )
        
        if response.status_code == 200:
            return response.json()['id']
        else:
            raise Exception(f"Failed to create CMMI page: {response.text}")
    
    def export_all(self,
                   sprint_metadata: Dict,
                   staff_metrics: pd.DataFrame,
                   team_metrics: pd.DataFrame,
                   cmmi_measures: Dict) -> Dict:
        """
        Export all sprint data to Notion
        
        Returns:
            Dict with created page IDs
        """
        print("Creating Sprint Summary page...")
        sprint_id = self.create_sprint_summary_page(sprint_metadata, cmmi_measures)
        
        print("Creating Staff Performance pages...")
        staff_pages = self.create_staff_performance_pages(staff_metrics, sprint_id)
        
        print("Creating Team Performance pages...")
        team_pages = self.create_team_performance_pages(team_metrics, sprint_id)
        
        print("Creating CMMI History page...")
        cmmi_page = self.create_cmmi_history_page(cmmi_measures, sprint_id)
        
        return {
            'sprint_id': sprint_id,
            'staff_pages': staff_pages,
            'team_pages': team_pages,
            'cmmi_page': cmmi_page
        }
    
    def export_to_json(self,
                       sprint_metadata: Dict,
                       staff_metrics: pd.DataFrame,
                       team_metrics: pd.DataFrame,
                       cmmi_measures: Dict,
                       output_path: str):
        """
        Export data to JSON files (alternative to Notion API)
        """
        data = {
            'sprint_metadata': {
                'sprint_name': sprint_metadata['sprint_name'],
                'date_from': str(sprint_metadata['date_from']),
                'date_to': str(sprint_metadata['date_to']),
                'product': sprint_metadata['product']
            },
            'staff_metrics': staff_metrics.to_dict(orient='records'),
            'team_metrics': team_metrics.to_dict(orient='records'),
            'cmmi_measures': cmmi_measures
        }
        
        with open(output_path, 'w') as f:
            json.dump(data, f, indent=2, default=str)
        
        print(f"Data exported to {output_path}")
    
    def export_to_csv(self,
                      staff_metrics: pd.DataFrame,
                      team_metrics: pd.DataFrame,
                      output_prefix: str):
        """
        Export metrics to CSV files
        """
        staff_metrics.to_csv(f"{output_prefix}_staff_metrics.csv", index=False)
        team_metrics.to_csv(f"{output_prefix}_team_metrics.csv", index=False)
        
        print(f"CSVs exported to {output_prefix}_*.csv")


def main():
    """Example usage"""
    
    # Example configuration
    database_ids = {
        'sprint_summary': 'YOUR_SPRINT_SUMMARY_DATABASE_ID',
        'staff_performance': 'YOUR_STAFF_PERFORMANCE_DATABASE_ID',
        'team_performance': 'YOUR_TEAM_PERFORMANCE_DATABASE_ID',
        'cmmi_history': 'YOUR_CMMI_HISTORY_DATABASE_ID'
    }
    
    # Example data
    sprint_metadata = {
        'sprint_name': 'Check Now 26.17',
        'date_from': datetime(2026, 1, 25),
        'date_to': datetime(2026, 2, 5),
        'product': 'Check Now'
    }
    
    staff_metrics = pd.DataFrame([
        {
            'staff_name': 'Ahmed Talaat',
            'team': 'Development',
            'kpi': 1.0229,
            'performance_rate': 0.9884,
            'utilization': 1.0159,
            'done_tasks_pct': 1.0000,
            'midsprint_addition_pct': 0.0323,
            'adhoc_pct': 0.0161
        }
    ])
    
    team_metrics = pd.DataFrame([
        {
            'team': 'Development',
            'kpi': 0.9234,
            'performance_rate': 0.8956,
            'utilization': 1.1234,
            'done_tasks_pct': 0.9456
        }
    ])
    
    cmmi_measures = {
        'completion_rate': 1.79,
        'effort_estimation_accuracy': 0.85,
        'bug_fixing_effort_pct': 0.0026,
        'utilization_rate': 1.29,
        'cmmi_productivity': 1.39
    }
    
    # Export to JSON (doesn't require Notion API)
    exporter = NotionExporter(api_key='dummy_key', database_ids=database_ids)
    exporter.export_to_json(
        sprint_metadata,
        staff_metrics,
        team_metrics,
        cmmi_measures,
        '/home/claude/sprint_export.json'
    )
    
    # Export to CSV
    exporter.export_to_csv(
        staff_metrics,
        team_metrics,
        '/home/claude/sprint_26_17'
    )


if __name__ == "__main__":
    main()
