"""
Sprint Analysis Automation - Calculator Module

This module contains all KPI and metric calculation logic.
"""

import pandas as pd
import numpy as np
from typing import Dict, List


class SprintCalculator:
    """Calculate all sprint metrics and KPIs"""
    
    def __init__(self, kpi_weights: Dict = None):
        """
        Initialize calculator with KPI weights
        
        Args:
            kpi_weights: Dict with 'performance_rate', 'utilization', 'done_tasks' weights
        """
        self.kpi_weights = kpi_weights or {
            'performance_rate': 0.2,
            'utilization': 0.1,
            'done_tasks': 0.7
        }
    
    def calculate_performance_rate(self, 
                                   original_estimate: float, 
                                   completed_work: float, 
                                   remaining_work: float) -> float:
        """
        Calculate Performance Rate
        
        Formula: Original Estimate / (Completed Work + Remaining Work)
        
        Returns:
            Performance rate (higher is better, ~1.0 is ideal)
        """
        denominator = completed_work + remaining_work
        if denominator == 0:
            return 0.0
        return original_estimate / denominator
    
    def calculate_utilization(self, 
                             completed_work: float, 
                             capacity: float) -> float:
        """
        Calculate Utilization
        
        Formula: Completed Work / Capacity
        
        Returns:
            Utilization rate (1.0 = 100% utilized)
        """
        if capacity == 0:
            return 0.0
        return completed_work / capacity
    
    def calculate_done_tasks_percentage(self, 
                                        done_count: int, 
                                        total_count: int) -> float:
        """
        Calculate Done Tasks Percentage
        
        Formula: COUNT(Done Tasks) / COUNT(All Tasks)
        
        Returns:
            Percentage of tasks completed (0.0 to 1.0)
        """
        if total_count == 0:
            return 0.0
        return done_count / total_count
    
    def calculate_midsprint_addition_percentage(self,
                                                midsprint_work: float,
                                                total_work: float) -> float:
        """
        Calculate MidSprint Addition Percentage
        
        Formula: SUM(Work with MidSprint tag) / SUM(All Work)
        
        Returns:
            Percentage of work that was mid-sprint addition (0.0 to 1.0)
        """
        if total_work == 0:
            return 0.0
        return midsprint_work / total_work
    
    def calculate_adhoc_percentage(self,
                                   adhoc_work: float,
                                   total_work: float) -> float:
        """
        Calculate Ad-hoc Percentage
        
        Formula: SUM(Work with Ad-hoc tag) / SUM(All Work)
        
        Returns:
            Percentage of work that was ad-hoc (0.0 to 1.0)
        """
        if total_work == 0:
            return 0.0
        return adhoc_work / total_work
    
    def calculate_kpi(self,
                     performance_rate: float,
                     utilization: float,
                     done_tasks_pct: float) -> float:
        """
        Calculate composite KPI
        
        Formula: (Performance Rate × W1) + (Utilization × W2) + (Done Tasks % × W3)
        
        Returns:
            Composite KPI score
        """
        kpi = (
            performance_rate * self.kpi_weights['performance_rate'] +
            utilization * self.kpi_weights['utilization'] +
            done_tasks_pct * self.kpi_weights['done_tasks']
        )
        return kpi
    
    def calculate_staff_metrics(self,
                                staff_agg: pd.DataFrame,
                                capacity_data: pd.DataFrame,
                                azure_data: pd.DataFrame) -> pd.DataFrame:
        """
        Calculate all metrics for each staff member
        
        Args:
            staff_agg: Aggregated staff data
            capacity_data: Capacity data by staff
            azure_data: Raw Azure DevOps data
        
        Returns:
            DataFrame with all calculated metrics
        """
        results = []
        
        for _, staff in staff_agg.iterrows():
            # Get capacity
            capacity_row = capacity_data[
                capacity_data['staff_member'] == staff['staff_name'] + ' <' + staff['staff_email'] + '>'
            ]
            
            if len(capacity_row) == 0:
                capacity = 63  # Default capacity
            else:
                capacity = capacity_row.iloc[0]['remaining_capacity']
            
            # Calculate MidSprint and Ad-hoc work
            staff_tasks = azure_data[azure_data['staff_email'] == staff['staff_email']]
            midsprint_work = staff_tasks[staff_tasks['has_midsprint_addition']]['Completed Work'].sum()
            adhoc_work = staff_tasks[staff_tasks['has_adhoc']]['Completed Work'].sum()
            
            # Calculate metrics
            performance_rate = self.calculate_performance_rate(
                staff['original_estimate_total'],
                staff['completed_work_total'],
                staff['remaining_work_total']
            )
            
            utilization = self.calculate_utilization(
                staff['completed_work_total'],
                capacity
            )
            
            done_tasks_pct = self.calculate_done_tasks_percentage(
                staff['done_count'],
                staff['total_count']
            )
            
            midsprint_pct = self.calculate_midsprint_addition_percentage(
                midsprint_work,
                staff['completed_work_total']
            )
            
            adhoc_pct = self.calculate_adhoc_percentage(
                adhoc_work,
                staff['completed_work_total']
            )
            
            kpi = self.calculate_kpi(
                performance_rate,
                utilization,
                done_tasks_pct
            )
            
            results.append({
                'staff_name': staff['staff_name'],
                'staff_email': staff['staff_email'],
                'team': staff['team'],
                'original_estimate': staff['original_estimate_total'],
                'completed_work': staff['completed_work_total'],
                'remaining_work': staff['remaining_work_total'],
                'capacity': capacity,
                'done_count': staff['done_count'],
                'total_count': staff['total_count'],
                'performance_rate': performance_rate,
                'utilization': utilization,
                'done_tasks_pct': done_tasks_pct,
                'midsprint_addition_pct': midsprint_pct,
                'adhoc_pct': adhoc_pct,
                'kpi': kpi
            })
        
        return pd.DataFrame(results)
    
    def calculate_team_metrics(self,
                               team_agg: pd.DataFrame,
                               capacity_data: pd.DataFrame,
                               azure_data: pd.DataFrame) -> pd.DataFrame:
        """
        Calculate all metrics for each team
        
        Args:
            team_agg: Aggregated team data
            capacity_data: Capacity data
            azure_data: Raw Azure DevOps data
        
        Returns:
            DataFrame with all calculated team metrics
        """
        results = []
        
        for _, team in team_agg.iterrows():
            # Get team capacity
            team_capacity = capacity_data[
                capacity_data['team'] == team['team']
            ]['remaining_capacity'].sum()
            
            # Calculate MidSprint and Ad-hoc work
            team_tasks = azure_data[azure_data['team'] == team['team']]
            midsprint_work = team_tasks[team_tasks['has_midsprint_addition']]['Completed Work'].sum()
            adhoc_work = team_tasks[team_tasks['has_adhoc']]['Completed Work'].sum()
            
            # Calculate metrics
            performance_rate = self.calculate_performance_rate(
                team['original_estimate_total'],
                team['completed_work_total'],
                team['remaining_work_total']
            )
            
            utilization = self.calculate_utilization(
                team['completed_work_total'],
                team_capacity
            )
            
            done_tasks_pct = self.calculate_done_tasks_percentage(
                team['done_count'],
                team['total_count']
            )
            
            midsprint_pct = self.calculate_midsprint_addition_percentage(
                midsprint_work,
                team['completed_work_total']
            )
            
            adhoc_pct = self.calculate_adhoc_percentage(
                adhoc_work,
                team['completed_work_total']
            )
            
            kpi = self.calculate_kpi(
                performance_rate,
                utilization,
                done_tasks_pct
            )
            
            results.append({
                'team': team['team'],
                'original_estimate': team['original_estimate_total'],
                'completed_work': team['completed_work_total'],
                'remaining_work': team['remaining_work_total'],
                'capacity': team_capacity,
                'done_count': team['done_count'],
                'total_count': team['total_count'],
                'performance_rate': performance_rate,
                'utilization': utilization,
                'done_tasks_pct': done_tasks_pct,
                'midsprint_addition_pct': midsprint_pct,
                'adhoc_pct': adhoc_pct,
                'kpi': kpi
            })
        
        return pd.DataFrame(results)
    
    def calculate_cmmi_measures(self,
                                sprint_metadata: Dict,
                                azure_data: pd.DataFrame) -> Dict:
        """
        Calculate CMMI measures
        
        Args:
            sprint_metadata: Sprint metadata from Data sheet
            azure_data: Raw Azure DevOps data
        
        Returns:
            Dict with CMMI measures
        """
        cmmi = {}
        
        # 1. Completion Rate
        cmmi['completion_rate'] = (
            sprint_metadata['actual_sp'] / sprint_metadata['planned_sp']
            if sprint_metadata['planned_sp'] > 0 else 0
        )
        cmmi['completion_rate_analysis'] = (
            'Good' if cmmi['completion_rate'] >= 0.90 else 'Needs Improvement'
        )
        
        # 2. Effort Estimation Accuracy
        cmmi['effort_estimation_accuracy'] = (
            sprint_metadata['estimated_dev_effort'] / sprint_metadata['actual_dev_effort']
            if sprint_metadata['actual_dev_effort'] > 0 else 0
        )
        cmmi['effort_estimation_analysis'] = (
            'Good' if 0.90 <= cmmi['effort_estimation_accuracy'] <= 1.20 
            else 'Needs Improvement'
        )
        
        # 3. Bug Fixing Effort Percentage
        cmmi['bug_fixing_effort_pct'] = (
            sprint_metadata['bug_fixing_effort'] / sprint_metadata['actual_dev_effort']
            if sprint_metadata['actual_dev_effort'] > 0 else 0
        )
        
        # 4. Utilization Rate
        cmmi['utilization_rate'] = (
            sprint_metadata['sum_actual_hours'] / sprint_metadata['sum_working_hours']
            if sprint_metadata['sum_working_hours'] > 0 else 0
        )
        
        # 5. CMMI Productivity
        if cmmi['utilization_rate'] > 0:
            cmmi['cmmi_productivity'] = (
                sprint_metadata['actual_sp'] / cmmi['utilization_rate']
            )
        else:
            cmmi['cmmi_productivity'] = 0
        
        # 6. Sprint Name
        cmmi['sprint_name'] = sprint_metadata['sprint_name']
        cmmi['date_from'] = sprint_metadata['date_from']
        cmmi['date_to'] = sprint_metadata['date_to']
        
        return cmmi
    
    def validate_metrics(self, metrics: pd.DataFrame) -> Dict:
        """
        Validate calculated metrics for anomalies
        
        Args:
            metrics: DataFrame with calculated metrics
        
        Returns:
            Dict with validation warnings
        """
        warnings = []
        
        # Check for extremely high utilization
        high_util = metrics[metrics['utilization'] > 1.5]
        if len(high_util) > 0:
            for _, row in high_util.iterrows():
                warnings.append({
                    'type': 'high_utilization',
                    'staff': row.get('staff_name', row.get('team')),
                    'value': row['utilization'],
                    'message': f"Utilization {row['utilization']:.2%} exceeds 150%"
                })
        
        # Check for very low performance rate
        low_perf = metrics[metrics['performance_rate'] < 0.5]
        if len(low_perf) > 0:
            for _, row in low_perf.iterrows():
                warnings.append({
                    'type': 'low_performance_rate',
                    'staff': row.get('staff_name', row.get('team')),
                    'value': row['performance_rate'],
                    'message': f"Performance rate {row['performance_rate']:.2f} is below 0.5"
                })
        
        # Check for KPI outliers
        kpi_outliers = metrics[(metrics['kpi'] < 0.3) | (metrics['kpi'] > 1.5)]
        if len(kpi_outliers) > 0:
            for _, row in kpi_outliers.iterrows():
                warnings.append({
                    'type': 'kpi_outlier',
                    'staff': row.get('staff_name', row.get('team')),
                    'value': row['kpi'],
                    'message': f"KPI {row['kpi']:.2f} is outside normal range (0.3-1.5)"
                })
        
        return {
            'status': 'warning' if warnings else 'success',
            'warnings': warnings
        }


def main():
    """Example usage"""
    # Example data
    staff_agg = pd.DataFrame([
        {
            'staff_name': 'Ahmed Talaat',
            'staff_email': 'atalaat@youxel.com',
            'team': 'Development',
            'original_estimate_total': 89,
            'completed_work_total': 64,
            'remaining_work_total': 1,
            'done_count': 10,
            'total_count': 10,
            'midsprint_count': 0,
            'adhoc_count': 1
        }
    ])
    
    capacity_data = pd.DataFrame([
        {
            'staff_member': 'Ahmed Talaat <atalaat@youxel.com>',
            'team': 'Development',
            'remaining_capacity': 63
        }
    ])
    
    azure_data = pd.DataFrame([
        {
            'staff_email': 'atalaat@youxel.com',
            'Completed Work': 64,
            'has_midsprint_addition': False,
            'has_adhoc': True,
            'team': 'Development'
        }
    ])
    
    # Initialize calculator
    calculator = SprintCalculator()
    
    # Calculate staff metrics
    staff_metrics = calculator.calculate_staff_metrics(
        staff_agg, capacity_data, azure_data
    )
    
    print("Staff Metrics:")
    print(staff_metrics)
    
    # Validate metrics
    validation = calculator.validate_metrics(staff_metrics)
    print("\nValidation:")
    print(validation)


if __name__ == "__main__":
    main()
