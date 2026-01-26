"""
BankVista - Dynamic Single-File Excel Dashboard
REVOLUTIONARY: One Excel file with dropdown for ALL branches
Save as: app.py
Run with: streamlit run app.py
"""

import streamlit as st
import pandas as pd
import plotly.graph_objects as go
from datetime import date
import io
import zipfile
from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter
from openpyxl.chart import BarChart, Reference
from openpyxl.worksheet.datavalidation import DataValidation

st.set_page_config(page_title="BankVista", page_icon="📊", layout="wide")

st.markdown("""<style>
.main-title {font-size: 2.5rem; font-weight: 700; color: #1a1a1a; margin-bottom: 0.5rem;}
.subtitle {font-size: 1.1rem; color: #666; margin-bottom: 2rem;}
.section-header {font-size: 1.4rem; font-weight: 600; color: #2c3e50; margin-top: 2rem; 
                 margin-bottom: 1rem; padding-bottom: 0.5rem; border-bottom: 2px solid #3498db;}
.alert-critical {background: #fff5f5; border-left: 4px solid #e74c3c; padding: 1rem; margin: 1rem 0; border-radius: 4px;}
.alert-warning {background: #fffef0; border-left: 4px solid #f39c12; padding: 1rem; margin: 1rem 0; border-radius: 4px;}
.alert-success {background: #f0fff4; border-left: 4px solid #27ae60; padding: 1rem; margin: 1rem 0; border-radius: 4px;}
.rec-item {background: #f8f9fa; padding: 0.8rem; margin: 0.5rem 0; border-radius: 4px; border-left: 3px solid #3498db;}
.feature-box {background: linear-gradient(135deg, #667eea 0%, #764ba2 100%); color: white; padding: 1.5rem; border-radius: 8px; margin: 1rem 0;}
#MainMenu {visibility: hidden;} footer {visibility: hidden;}
</style>""", unsafe_allow_html=True)

class EnhancedBankVista:
    def __init__(self, data):
        self.data = data
        self.insights = {'critical': [], 'warning': [], 'success': []}
        self.recommendations = []
        self.priorities = []
        
    def analyze(self):
        self._check_deposits()
        self._check_advances()
        self._check_npa()
        self._check_profit()
        self._check_casa()
        self._check_cd_ratio()
        self._check_staff_productivity()
        self._set_priorities()
        grade, score = self._calculate_grade()
        return {'grade': grade, 'score': score, 'insights': self.insights, 
                'recommendations': self.recommendations, 'priorities': self.priorities}
    
    def _check_deposits(self):
        total = float(self.data.get('Total_Deposits', 0))
        target = float(self.data.get('Deposit_Target', 1))
        pct = (total/target*100) if target>0 else 0
        gap = target-total
        
        if pct<85:
            self.insights['critical'].append({'title':'Deposits Severely Low','detail':f'{pct:.1f}% achievement (₹{abs(gap):.1f}Cr shortfall)'})
            self.recommendations.extend([
                f"📌 Mobilize ₹{abs(gap)/7:.2f}Cr daily for next 7 days",
                "🎯 Launch CASA-focused deposit campaign",
                "👥 Conduct customer meet programs",
                "💼 Target bulk deposits from corporates"
            ])
            self.priorities.append({'Priority':'P1','Area':'Deposits','Gap':f'₹{abs(gap):.1f}Cr','Action':'Urgent mobilization drive','Timeline':'7 days'})
        elif pct<95:
            self.insights['warning'].append({'title':'Deposits Below Target','detail':f'{pct:.1f}% (₹{abs(gap):.1f}Cr gap)'})
            self.recommendations.append(f"🎯 Mobilize ₹{abs(gap)/14:.2f}Cr daily for 14 days")
        else:
            self.insights['success'].append({'title':'Deposits Excellent','detail':f'{pct:.1f}% achievement'})
    
    def _check_advances(self):
        total = float(self.data.get('Advances', 0))
        target = float(self.data.get('Advance_Target', 1))
        pct = (total/target*100) if target>0 else 0
        gap = target-total
        
        if pct<85:
            self.insights['critical'].append({'title':'Advances Critical','detail':f'{pct:.1f}% (₹{abs(gap):.1f}Cr shortfall)'})
            self.recommendations.extend([
                f"💳 Disburse ₹{abs(gap)/7:.2f}Cr daily for 7 days",
                "🏭 Focus on MSME lending",
                "🏠 Home loan campaigns",
                "⚡ Fast-track sanctioned files"
            ])
            self.priorities.append({'Priority':'P1','Area':'Advances','Gap':f'₹{abs(gap):.1f}Cr','Action':'Aggressive disbursement','Timeline':'7 days'})
        elif pct<95:
            self.insights['warning'].append({'title':'Advances Below Target','detail':f'{pct:.1f}%'})
            self.recommendations.append(f"💳 Disburse ₹{abs(gap)/14:.2f}Cr daily")
        else:
            self.insights['success'].append({'title':'Advances Strong','detail':f'{pct:.1f}%'})
    
    def _check_npa(self):
        npa = float(self.data.get('NPA_Percent', 0))
        if npa>6:
            self.insights['critical'].append({'title':'NPA Critical','detail':f'{npa:.2f}% - Urgent action required'})
            self.recommendations.extend([
                "⚠️ Daily recovery task force meetings",
                "⚖️ Initiate legal action on top 10 defaulters",
                "💰 Launch OTS schemes",
                "📞 Call all overdue accounts daily"
            ])
            self.priorities.append({'Priority':'P1','Area':'NPA','Gap':f'{npa:.1f}%','Action':'Recovery war room','Timeline':'Immediate'})
        elif npa>3:
            self.insights['warning'].append({'title':'NPA Elevated','detail':f'{npa:.2f}% - Monitor closely'})
            self.recommendations.append("📊 Weekly recovery reviews")
        else:
            self.insights['success'].append({'title':'NPA Healthy','detail':f'{npa:.2f}%'})
    
    def _check_profit(self):
        profit = float(self.data.get('Profit_Per_Staff', 0))
        if profit<2:
            self.insights['critical'].append({'title':'Profitability Low','detail':f'₹{profit:.2f}L per staff'})
            self.recommendations.extend([
                "💰 Increase fee-based income",
                "🎯 Cross-sell insurance & mutual funds",
                "✂️ Optimize operating costs"
            ])
        elif profit<3:
            self.insights['warning'].append({'title':'Profit Below Par','detail':f'₹{profit:.2f}L'})
            self.recommendations.append("📈 Focus on high-margin products")
        else:
            self.insights['success'].append({'title':'Profitability Good','detail':f'₹{profit:.2f}L/staff'})
    
    def _check_casa(self):
        casa = float(self.data.get('CASA_Percent', 0))
        if casa<30:
            self.insights['warning'].append({'title':'CASA Low','detail':f'{casa:.1f}% (Target: 40%+)'})
            self.recommendations.extend([
                "💳 Launch salary account campaign",
                "🎁 CASA account opening incentives"
            ])
        elif casa>=40:
            self.insights['success'].append({'title':'CASA Excellent','detail':f'{casa:.1f}%'})
    
    def _check_cd_ratio(self):
        cd = float(self.data.get('CD_Ratio', 0))
        if cd>80:
            self.insights['warning'].append({'title':'CD Ratio High','detail':f'{cd:.1f}% - Liquidity pressure'})
            self.recommendations.append("💵 Priority: Deposit mobilization")
        elif cd<60:
            self.insights['warning'].append({'title':'CD Ratio Low','detail':f'{cd:.1f}% - Underutilized funds'})
            self.recommendations.append("📈 Increase lending activities")
        else:
            self.insights['success'].append({'title':'CD Ratio Optimal','detail':f'{cd:.1f}%'})
    
    def _check_staff_productivity(self):
        bus_staff = float(self.data.get('Business_Per_Staff', 0))
        if bus_staff<50:
            self.insights['warning'].append({'title':'Staff Productivity Low','detail':f'₹{bus_staff:.1f}Cr/staff'})
            self.recommendations.extend([
                "👨‍🏫 Staff training programs",
                "🎯 Individual performance targets"
            ])
        elif bus_staff>=80:
            self.insights['success'].append({'title':'Staff Highly Productive','detail':f'₹{bus_staff:.1f}Cr/staff'})
    
    def _set_priorities(self):
        if not self.priorities:
            self.priorities.append({'Priority':'P3','Area':'Overall','Gap':'-','Action':'Maintain performance','Timeline':'Ongoing'})
    
    def _calculate_grade(self):
        try:
            d_a = float(self.data.get('Total_Deposits',0))
            d_t = float(self.data.get('Deposit_Target',1))
            a_a = float(self.data.get('Advances',0))
            a_t = float(self.data.get('Advance_Target',1))
            npa = float(self.data.get('NPA_Percent',0))
            pft = float(self.data.get('Profit_Per_Staff',0))
            casa = float(self.data.get('CASA_Percent',0))
            cd = float(self.data.get('CD_Ratio',0))
            
            s1 = min((d_a/d_t*25),25) if d_t>0 else 0
            s2 = min((a_a/a_t*25),25) if a_t>0 else 0
            s3 = 20 if npa<=3 else (12 if npa<=6 else 5)
            s4 = 15 if pft>=5 else (10 if pft>=3 else 5)
            s5 = 10 if casa>=40 else (5 if casa>=30 else 2)
            s6 = 5 if 60<=cd<=80 else 2
            
            total = round(s1+s2+s3+s4+s5+s6,1)
            grade = "A+" if total>=90 else ("A" if total>=80 else ("B" if total>=65 else ("C" if total>=50 else "D")))
            return grade, total
        except:
            return "N/A", 0

def sample_data():
    return pd.DataFrame({
        'Branch_ID': ['B1001','B1002','B1003','B1004','B1005','B2001','B2002','B2003'],
        'Branch_Name': ['Mansoorabad','Adilabad','Hyderabad Main','Secunderabad','Warangal','Vijayawada','Visakhapatnam','Guntur'],
        'Zone': ['Telangana']*5 + ['Andhra Pradesh']*3,
        'Total_Deposits': [110.97,85.45,245.80,189.23,67.89,198.76,223.45,145.67],
        'Deposit_Target': [105.46,95.00,250.00,195.00,75.00,205.00,220.00,150.00],
        'Advances': [232.29,156.78,412.50,289.45,98.76,356.89,389.23,245.78],
        'Advance_Target': [218.16,175.00,425.00,295.00,110.00,365.00,395.00,255.00],
        'NPA_Percent': [2.8,5.4,1.9,3.2,8.5,2.3,1.8,4.1],
        'Profit_Per_Staff': [5.44,3.1,6.8,4.5,1.8,5.8,6.5,3.8],
        'CASA_Percent': [42.3,28.5,51.2,38.7,25.4,44.7,49.3,34.9],
        'CD_Ratio': [72.5,78.2,65.8,70.1,82.3,66.8,64.5,73.4],
        'Business_Per_Staff': [85.2,58.3,95.7,78.9,45.6,86.7,91.2,67.8],
        'Staff_Count': [25,18,42,35,15,38,40,27]
    })

def load_file(file):
    try:
        ext = file.name.split('.')[-1].lower()
        if ext=='csv':
            return pd.read_csv(file)
        elif ext in ['xlsx','xls']:
            return pd.read_excel(file)
        else:
            st.error("Unsupported format")
            return None
    except Exception as e:
        st.error(f"Error: {e}")
        return None

def create_dynamic_excel(df):
    """
    REVOLUTIONARY: Creates ONE Excel file with dropdown to select ANY branch
    All data and insights update automatically via Excel formulas
    """
    output = io.BytesIO()
    wb = Workbook()
    wb.remove(wb.active)
    
    thin_border = Border(left=Side(style='thin'), right=Side(style='thin'), 
                         top=Side(style='thin'), bottom=Side(style='thin'))
    
    # ========== DATA SHEET (Hidden) ==========
    ws_data = wb.create_sheet("_Data")
    ws_data.sheet_state = 'hidden'
    df = df.sort_values("Branch_Name").reset_index(drop=True)
    headers = df.columns.tolist()
    # Helper column for search-based dropdown
    ws_data.cell(1, len(headers) + 1, "SearchMatch").font = Font(bold=True)

    for r in range(2, len(df) + 2):
        ws_data.cell(
            r,
            len(headers) + 1,
            f'=IF(ISNUMBER(SEARCH(Dashboard!$B$3,B{r})),B{r},"")'
        )

    
    # Write all branch data

    for col_idx, header in enumerate(headers, 1):
        ws_data.cell(1, col_idx, header).font = Font(bold=True)
    
    for row_idx, row in enumerate(df.itertuples(index=False), 2):
        for col_idx, value in enumerate(row, 1):
            ws_data.cell(row_idx, col_idx, value)
    
    # ========== DASHBOARD SHEET (Main Interactive) ==========
    ws = wb.create_sheet("Dashboard", 0)
    
    # Title
    ws.merge_cells('A1:H1')
    ws['A1'] = "🤖 BANKVISTA - DYNAMIC DASHBOARD"
    ws['A1'].font = Font(bold=True, size=16, color="FFFFFF")
    ws['A1'].fill = PatternFill("solid", fgColor="1F4E78")
    ws['A1'].alignment = Alignment(horizontal='center', vertical='center')
    ws.row_dimensions[1].height = 30
    
    # Dropdown selector
    ws['A3'] = "Type Branch Name:"
    ws['A3'].font = Font(bold=True, size=12, color="1F4E78")
    ws.merge_cells('B3:D3')
    ws['B3'] = df.iloc[0]['Branch_Name']  # Default first branch
    ws['B3'].font = Font(size=12, bold=True)
    ws['B3'].fill = PatternFill("solid", fgColor="FFF2CC")
    ws['B3'].border = thin_border
    
    # Create dropdown validation
    # Dropdown based on range (enables typing first letter)
    last_row = len(df) + 1  # header + data
    search_col = get_column_letter(len(headers) + 1)

    dv = DataValidation(
        type="list",
        formula1=f"=_Data!${search_col}$2:${search_col}${last_row}",
        allow_blank=False
    )

    dv.add('B3')
    ws.add_data_validation(dv)

    ws['F3'] = "Date:"
    ws['F3'].font = Font(bold=True)
    ws['G3'] = date.today().strftime('%d-%b-%Y')
    
    # Lookup formulas (these update automatically when dropdown changes!)
    ws['A5'] = "Branch ID:"
    ws['B5'] = '=INDEX(_Data!$A:$A, MATCH(B3, _Data!$B:$B, 0))'
    ws['A6'] = "Zone:"
    ws['B6'] = '=INDEX(_Data!$C:$C, MATCH(B3, _Data!$B:$B, 0))'
    ws['B4'] = "💡 Type few letters, then select from list"
    ws['B4'].font = Font(italic=True, size=9, color="666666")    
    # Performance Cards with FORMULAS
    ws.merge_cells('A8:B8')
    ws['A8'] = "GRADE"
    ws['A8'].font = Font(bold=True, size=10, color="FFFFFF")
    ws['A8'].fill = PatternFill("solid", fgColor="27AE60")
    ws['A8'].alignment = Alignment(horizontal='center')
    ws.merge_cells('A9:B9')
    ws['A9'] = '=IF(C9>=90,"A+",IF(C9>=80,"A",IF(C9>=65,"B",IF(C9>=50,"C","D"))))'
    ws['A9'].font = Font(bold=True, size=16)
    ws['A9'].alignment = Alignment(horizontal='center')
    
    ws.merge_cells('D8:E8')
    ws['D8'] = "SCORE"
    ws['D8'].font = Font(bold=True, size=10, color="FFFFFF")
    ws['D8'].fill = PatternFill("solid", fgColor="3498DB")
    ws['D8'].alignment = Alignment(horizontal='center')
    ws.merge_cells('D9:E9')
    ws['D9'] = "=ROUND(C9,0)&\"/100\""
    ws['D9'].font = Font(bold=True, size=14)
    ws['D9'].alignment = Alignment(horizontal='center')
    
    ws.merge_cells('G8:H8')
    ws['G8'] = "STATUS"
    ws['G8'].font = Font(bold=True, size=10, color="FFFFFF")
    ws['G8'].fill = PatternFill("solid", fgColor="E67E22")
    ws['G8'].alignment = Alignment(horizontal='center')
    ws.merge_cells('G9:H9')
    ws['G9'] = '=IF(C9>=80,"✅ Excellent",IF(C9>=65,"⚠️ Good","🔴 Review"))'
    ws['G9'].font = Font(bold=True, size=12)
    ws['G9'].alignment = Alignment(horizontal='center')
    
    # Hidden score calculation cell
    ws['C9'] = '''=
    MIN(
     INDEX(_Data!$D:$D, MATCH(B3,_Data!$B:$B,0)) /
     INDEX(_Data!$E:$E, MATCH(B3,_Data!$B:$B,0)) * 25, 25
    )
    +
    MIN(
     INDEX(_Data!$F:$F, MATCH(B3,_Data!$B:$B,0)) /
     INDEX(_Data!$G:$G, MATCH(B3,_Data!$B:$B,0)) * 25, 25
    )
    +
    IF(INDEX(_Data!$H:$H, MATCH(B3,_Data!$B:$B,0))<=3,20,
     IF(INDEX(_Data!$H:$H, MATCH(B3,_Data!$B:$B,0))<=6,12,5)
    )
    +
    IF(INDEX(_Data!$I:$I, MATCH(B3,_Data!$B:$B,0))>=5,15,
     IF(INDEX(_Data!$I:$I, MATCH(B3,_Data!$B:$B,0))>=3,10,5)
    )
    +
    IF(INDEX(_Data!$J:$J, MATCH(B3,_Data!$B:$B,0))>=40,10,
     IF(INDEX(_Data!$J:$J, MATCH(B3,_Data!$B:$B,0))>=30,5,2)
    )
    +
    IF(AND(
     INDEX(_Data!$K:$K, MATCH(B3,_Data!$B:$B,0))>=60,
     INDEX(_Data!$K:$K, MATCH(B3,_Data!$B:$B,0))<=80
    ),5,2)
    '''

    ws['C9'].font = Font(size=8, color="FFFFFF")
    
    # Metrics Table with FORMULAS
    ws.merge_cells('A11:H11')
    ws['A11'] = "KEY FINANCIAL METRICS"
    ws['A11'].font = Font(bold=True, size=12, color="FFFFFF")
    ws['A11'].fill = PatternFill("solid", fgColor="4472C4")
    ws['A11'].alignment = Alignment(horizontal='center')
    
    headers_metrics = ['Metric','Actual','Target','Gap','Achievement %','Status']
    for i, h in enumerate(headers_metrics, 1):
        if i <= 6:
            c = ws.cell(12, i)
            c.value = h
            c.font = Font(bold=True, size=10, color="FFFFFF")
            c.fill = PatternFill("solid", fgColor="5B9BD5")
            c.alignment = Alignment(horizontal='center')
            c.border = thin_border
    
    metrics = [
        ("Deposits (Cr)", 4, 5),
        ("Advances (Cr)", 6, 7),
        ("NPA %", 8, None),
        ("Profit/Staff (L)", 9, None),
        ("CASA %", 10, None),
        ("CD Ratio %", 11, None)
    ]
    
    row = 13
    for metric, actual_col, target_col in metrics:
        ws.cell(row, 1, metric)
        ws.cell(
            row, 2,
            f'=INDEX(_Data!${get_column_letter(actual_col)}:${get_column_letter(actual_col)}, MATCH(B3,_Data!$B:$B,0))'
        )

        ws.cell(row, 2).number_format = '#,##0.00'
        
        if target_col:
            ws.cell(
                row, 3,
                f'=INDEX(_Data!${get_column_letter(target_col)}:${get_column_letter(target_col)}, MATCH(B3,_Data!$B:$B,0))'
        )

            ws.cell(row, 3).number_format = '#,##0.00'
            ws.cell(row, 4, f'=B{row}-C{row}')
            ws.cell(row, 4).number_format = '#,##0.00'
            ws.cell(row, 5, f'=B{row}/C{row}')
            ws.cell(row, 5).number_format = '0.0%'
            ws.cell(row, 6, f'=IF(B{row}>=C{row},"✅ On Track","🔴 Gap")')
        else:
            if metric == "NPA %":
                ws.cell(row, 3, "3.00%")
                ws.cell(row, 6, f'=IF(B{row}<=3,"✅ Good","🔴 High")')
            elif metric == "Profit/Staff (L)":
                ws.cell(row, 3, "5.00")
                ws.cell(row, 6, f'=IF(B{row}>=5,"✅ Excellent",IF(B{row}>=3,"✅ Good","🔴 Low"))')
            elif metric == "CASA %":
                ws.cell(row, 3, "40.00%")
                ws.cell(row, 6, f'=IF(B{row}>=40,"✅ Excellent","🔴 Low")')
            else:  # CD Ratio
                ws.cell(row, 3, "70.00%")
                ws.cell(row, 6, f'=IF(AND(B{row}>=60,B{row}<=80),"✅ Optimal","⚠️ Review")')
        
        for col in range(1, 7):
            ws.cell(row, col).border = thin_border
            ws.cell(row, col).alignment = Alignment(horizontal='center')
        
        row += 1
    
    # Instructions
    ws.merge_cells('A20:H20')
    ws['A20'] = "💡 INSTRUCTIONS"
    ws['A20'].font = Font(bold=True, size=12, color="FFFFFF")
    ws['A20'].fill = PatternFill("solid", fgColor="27AE60")
    ws['A20'].alignment = Alignment(horizontal='center')
    
    instructions = [
        "1. Click on cell B3 dropdown to select any branch",
        "2. All metrics, grades, and insights update AUTOMATICALLY",
        "3. No internet required - works completely offline",
        "4. Share this single file across entire organization",
        "5. Perfect for reviews, meetings, and audits"
    ]
    
    for idx, inst in enumerate(instructions, 21):
        ws.merge_cells(f'A{idx}:H{idx}')
        ws[f'A{idx}'] = inst
        ws[f'A{idx}'].alignment = Alignment(wrap_text=True)
        ws.row_dimensions[idx].height = 20
    
    # Set widths
    ws.column_dimensions['A'].width = 20
    for c in ['B','C','D','E','F','G','H']:
        ws.column_dimensions[c].width = 15
    
    # Summary Sheet
    ws_summary = wb.create_sheet("All Branches Summary")
    ws_summary.merge_cells('A1:H1')
    ws_summary['A1'] = "ALL BRANCHES PERFORMANCE SUMMARY"
    ws_summary['A1'].font = Font(bold=True, size=14, color="FFFFFF")
    ws_summary['A1'].fill = PatternFill("solid", fgColor="1F4E78")
    ws_summary['A1'].alignment = Alignment(horizontal='center')
    
    summary_headers = ['Branch', 'Zone', 'Deposits %', 'Advances %', 'NPA %', 'CASA %', 'Grade', 'Status']
    for i, h in enumerate(summary_headers, 1):
        c = ws_summary.cell(3, i)
        c.value = h
        c.font = Font(bold=True, color="FFFFFF")
        c.fill = PatternFill("solid", fgColor="4472C4")
        c.alignment = Alignment(horizontal='center')
        c.border = thin_border
    
    for row_idx, row in df.iterrows():
        ai = EnhancedBankVista(row.to_dict())
        results = ai.analyze()
        
        r = row_idx + 4
        ws_summary.cell(r, 1, row['Branch_Name'])
        ws_summary.cell(r, 2, row['Zone'])
        ws_summary.cell(r, 3, f"{(row['Total_Deposits']/row['Deposit_Target']*100):.1f}%")
        ws_summary.cell(r, 4, f"{(row['Advances']/row['Advance_Target']*100):.1f}%")
        ws_summary.cell(r, 5, f"{row['NPA_Percent']:.2f}%")
        ws_summary.cell(r, 6, f"{row['CASA_Percent']:.1f}%")
        ws_summary.cell(r, 7, results['grade'])
        ws_summary.cell(r, 8, "✅" if results['grade'] in ['A+','A'] else ("⚠️" if results['grade']=='B' else "🔴"))
        
        for col in range(1, 9):
            ws_summary.cell(r, col).border = thin_border
            ws_summary.cell(r, col).alignment = Alignment(horizontal='center')
    
    for i in range(1, 9):
        ws_summary.column_dimensions[get_column_letter(i)].width = 18
    
    # User Guide Sheet
    ws_guide = wb.create_sheet("User Guide")
    ws_guide['A1'] = "🚀 BANKVISTA - USER GUIDE"
    ws_guide['A1'].font = Font(bold=True, size=16, color="FFFFFF")
    ws_guide['A1'].fill = PatternFill("solid", fgColor="667EEA")
    ws_guide.merge_cells('A1:D1')
    ws_guide['A1'].alignment = Alignment(horizontal='center')
    ws_guide.row_dimensions[1].height = 30
    
    guide_content = [
        ("", ""),
        ("WHAT MAKES THIS SPECIAL?", ""),
        ("✅ Single File Solution", "One Excel file for ALL branches - no more multiple files!"),
        ("✅ Dynamic Dropdown", "Select any branch and see instant updates"),
        ("✅ Automatic Calculations", "All metrics calculated via Excel formulas"),
        ("✅ Offline Ready", "Works without internet - perfect for intranet"),
        ("✅ AI-Powered Grading", "Smart scoring system (A+ to D grades)"),
        ("", ""),
        ("HOW TO USE:", ""),
        ("Step 1", "Go to 'Dashboard' sheet"),
        ("Step 2", "Click dropdown in cell B3"),
        ("Step 3", "Select any branch name"),
        ("Step 4", "Watch all data update automatically!"),
        ("", ""),
        ("FEATURES:", ""),
        ("📊 Real-time Metrics", "Deposits, Advances, NPA, CASA, CD Ratio"),
        ("🎯 Performance Grading", "A+ to D with automatic color coding"),
        ("📈 All Branches Summary", "See complete organization view"),
        ("💾 Fully Offline", "No internet needed - share freely"),
        ("", ""),
        ("BENEFITS:", ""),
        ("For Branch Managers", "Quick performance snapshot"),
        ("For Regional Heads", "Compare all branches instantly"),
        ("For Head Office", "Organization-wide visibility"),
        ("For Auditors", "Complete data in one place"),
    ]
    
    row = 3
    for title, desc in guide_content:
        if title:
            ws_guide[f'A{row}'] = title
            ws_guide[f'A{row}'].font = Font(bold=True, size=11)
            ws_guide[f'B{row}'] = desc
            ws_guide[f'B{row}'].alignment = Alignment(wrap_text=True)
        ws_guide.row_dimensions[row].height = 25
        row += 1
    
    ws_guide.column_dimensions['A'].width = 25
    ws_guide.column_dimensions['B'].width = 50
    
    wb.save(output)
    output.seek(0)
    return output

def main():
    st.markdown('<p class="main-title">🤖 BankVista - Next Generation Analytics</p>', unsafe_allow_html=True)
    st.markdown('<p class="subtitle">Revolutionary Single-File Dynamic Dashboard | Complete Offline Solution</p>', unsafe_allow_html=True)
    
    # Feature highlight
    st.markdown("""<div class="feature-box">
    <h3 style="margin-top:0;">🌟 REVOLUTIONARY FEATURE</h3>
    <p style="font-size:1.1rem;margin-bottom:0;">One Excel file with dropdown for ALL branches. Select any branch, everything updates automatically via Excel formulas. No macros, no internet needed!</p>
    </div>""", unsafe_allow_html=True)
    
    with st.sidebar:
        st.markdown("### 📤 Upload Data")
        st.info("✅ CSV, Excel\n💻 Works offline!\n🌐 Intranet ready")
        
        uploaded = st.file_uploader("Choose file", type=['csv','xlsx','xls'], key="uploader")
        st.markdown("---")
        sample_btn = st.button("📊 Try Sample Data", use_container_width=True)
        if sample_btn:
            st.session_state['use_sample'] = True
            st.session_state['uploaded_df'] = sample_data()
        
        st.markdown("---")
        st.success("**🎯 Key Features:**\n✅ Single dynamic Excel\n✅ Dropdown selection\n✅ Auto-calculations\n✅ Offline ready\n✅ No macros needed")
        
        st.markdown("---")
        st.markdown("### 🏆 Market Comparison")
        st.info("**vs Tableau:** Offline ✅\n**vs Power BI:** No license ✅\n**vs SAP:** Simple ✅\n**BankVista:** All in one ✅")
    
    # Load data
    if uploaded:
        df = load_file(uploaded)
        if df is not None:
            st.session_state['uploaded_df'] = df
            st.session_state['use_sample'] = False
    
    if st.session_state.get('use_sample') or 'uploaded_df' in st.session_state:
        df = st.session_state['uploaded_df']
        
        st.success(f"✅ Loaded {len(df)} branches")
        
        # Show preview
        selected = st.selectbox("🏢 Select Branch (Preview)", df['Branch_Name'].tolist(), key=f"branch_selector_{len(df)}")
        
        branch_data = df[df['Branch_Name']==selected].iloc[0].to_dict()
        ai = EnhancedBankVista(branch_data)
        results = ai.analyze()
        
        st.markdown("---")
        st.markdown('<p class="section-header">📊 Performance Preview</p>', unsafe_allow_html=True)
        
        c1,c2,c3,c4 = st.columns(4)
        grade_color = "🟢" if results['grade'] in ['A+','A'] else ("🟡" if results['grade']=='B' else "🔴")
        c1.metric("Grade", f"{grade_color} {results['grade']}")
        c2.metric("Score", f"{results['score']}/100")
        c3.metric("Critical", len(results['insights']['critical']))
        c4.metric("Actions", len(results['recommendations']))
        
        # Quick metrics
        col1, col2, col3, col4 = st.columns(4)
        npa = branch_data.get('NPA_Percent',0)
        col1.metric("NPA %", f"{npa:.2f}%", delta="Good" if npa<=3 else "High", delta_color="normal" if npa<=3 else "inverse")
        pft = branch_data.get('Profit_Per_Staff',0)
        col2.metric("Profit/Staff", f"₹{pft:.2f}L")
        casa = branch_data.get('CASA_Percent',0)
        col3.metric("CASA %", f"{casa:.1f}%")
        cd = branch_data.get('CD_Ratio',0)
        col4.metric("CD Ratio", f"{cd:.1f}%")
        
        st.markdown("---")
        st.markdown('<p class="section-header">📥 Download Dynamic Dashboard</p>', unsafe_allow_html=True)
        
        col1, col2 = st.columns([2,1])
        
        with col1:
            st.markdown("""
            ### 🎯 Dynamic Single-File Dashboard
            
            **What you'll get:**
            - ✅ **ONE Excel file** for ALL {count} branches
            - 🎯 **Dropdown selector** in cell B3
            - ⚡ **Auto-updating metrics** via Excel formulas
            - 📊 **Summary sheet** with all branches
            - 📖 **User guide** included
            - 💾 **Completely offline** - no macros, no internet
            
            **Perfect for:**
            - Branch managers reviewing their performance
            - Regional heads comparing branches
            - Head office monitoring all branches
            - Audit teams needing complete data
            - Management meetings and reviews
            """.format(count=len(df)))
            
            if st.button("🚀 Generate Dynamic Dashboard", use_container_width=True, type="primary"):
                with st.spinner(f"Creating dynamic Excel with {len(df)} branches..."):
                    dynamic_excel = create_dynamic_excel(df)
                
                st.success("✅ Dynamic dashboard created!")
                st.download_button(
                    "📊 Download Dynamic Dashboard (All Branches)", 
                    dynamic_excel,
                    f"BankVista_Dynamic_Dashboard_{date.today().strftime('%Y%m%d')}.xlsx",
                    "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                    use_container_width=True,
                    type="primary"
                )
        
        with col2:
            st.image("https://via.placeholder.com/300x400/667eea/ffffff?text=Single+File%0AAll+Branches%0ADropdown+Magic", use_container_width=True)
            st.caption("One file. All branches. Dynamic updates.")
        
        # Market Analysis
        st.markdown("---")
        st.markdown('<p class="section-header">🏆 Market Analysis & Competitive Edge</p>', unsafe_allow_html=True)
        
        col1, col2 = st.columns(2)
        
        with col1:
            st.markdown("""
            ### 💼 Existing Market Solutions
            
            **1. Tableau**
            - ❌ Expensive licenses ($70/user/month)
            - ❌ Requires internet for dashboards
            - ❌ Complex setup
            - ✅ Good visualizations
            
            **2. Power BI**
            - ❌ Microsoft ecosystem dependency
            - ❌ $10-20/user/month
            - ❌ Requires Power BI Desktop
            - ✅ Good for large enterprises
            
            **3. SAP Analytics Cloud**
            - ❌ Very expensive
            - ❌ Complex implementation
            - ❌ Requires training
            - ✅ Enterprise-grade
            
            **4. Qlik Sense**
            - ❌ Complex licensing
            - ❌ Steep learning curve
            - ✅ Good for advanced analytics
            
            **5. Excel Manual Reports**
            - ✅ Familiar to everyone
            - ❌ Multiple files chaos
            - ❌ Manual updates needed
            - ❌ Error-prone
            """)
        
        with col2:
            st.markdown("""
            ### 🚀 BankVista Advantages
            
            **Cost:**
            - ✅ **FREE** - No licensing fees
            - ✅ No per-user costs
            - ✅ No subscription needed
            
            **Accessibility:**
            - ✅ **Works offline** - Perfect for intranet
            - ✅ Just Excel - Everyone has it
            - ✅ No special software needed
            - ✅ Runs on any computer
            
            **Simplicity:**
            - ✅ **Dropdown + Auto-update** - That's it!
            - ✅ No training needed
            - ✅ Instant deployment
            - ✅ Share via email/network
            
            **Innovation:**
            - ✅ **AI-powered grading** system
            - ✅ Automatic insights
            - ✅ Smart recommendations
            - ✅ Priority matrix
            
            **Unique Features:**
            - 🌟 **Single file, all branches**
            - 🌟 **Excel formulas (no macros)**
            - 🌟 **Offline-first design**
            - 🌟 **Banking-specific logic**
            - 🌟 **Instant updates**
            """)
        
        st.markdown("---")
        st.markdown("### 💡 Suggestions for Market Leadership")
        
        tab1, tab2, tab3 = st.tabs(["🎯 Feature Enhancements", "📈 Monetization", "🚀 Go-to-Market"])
        
        with tab1:
            st.markdown("""
            ### Feature Roadmap
            
            **Phase 1: Current (✅ Done)**
            - Dynamic single-file dashboard
            - AI-powered insights
            - Offline capability
            
            **Phase 2: Next 30 Days**
            1. **Trend Analysis**
               - Month-over-month comparisons
               - Quarterly trends
               - Year-over-year growth
            
            2. **Peer Benchmarking**
               - Compare to zone average
               - Rank branches
               - Best practices from top performers
            
            3. **Predictive Alerts**
               - NPA early warning (7-day forecast)
               - Target miss prediction
               - Risk scoring
            
            **Phase 3: Next 90 Days**
            4. **Custom Dashboards**
               - RBI compliance view
               - Audit-ready reports
               - Board presentation mode
            
            5. **Multi-language Support**
               - Hindi, Tamil, Telugu, etc.
               - Regional language insights
            
            6. **Mobile-Friendly Export**
               - WhatsApp-ready format
               - PDF summary cards
            
            **Phase 4: Advanced**
            7. **Integration APIs**
               - CBS (Core Banking System) integration
               - Auto-import from Finacle/Flexcube
               - Email automation
            
            8. **ML-Powered Features**
               - Customer churn prediction
               - Cross-sell recommendations
               - Fraud detection alerts
            """)
        
        with tab2:
            st.markdown("""
            ### 💰 Monetization Strategy
            
            **Freemium Model:**
            - **Free Tier:** Up to 50 branches
            - **Pro Tier:** Unlimited branches + advanced features
            - **Enterprise:** White-label + custom features
            
            **Pricing Suggestions:**
            1. **BankVista Free**
               - Up to 50 branches
               - Basic dashboard
               - Community support
               - **Price: ₹0**
            
            2. **BankVista Pro**
               - Unlimited branches
               - Trend analysis
               - Predictive alerts
               - Email support
               - **Price: ₹9,999/month for organization**
            
            3. **BankVista Enterprise**
               - Everything in Pro
               - Custom branding
               - API integration
               - CBS integration
               - Dedicated support
               - Training included
               - **Price: ₹49,999/month**
            
            4. **One-Time Implementation**
               - Custom setup
               - Data migration
               - Training (5 sessions)
               - 3 months support
               - **Price: ₹1,99,999 one-time**
            
            **Revenue Potential:**
            - Target: 100 banks/NBFCs in Year 1
            - Mix: 60% Free, 30% Pro, 10% Enterprise
            - Revenue: ₹1.5-2 Cr ARR
            """)
        
        with tab3:
            st.markdown("""
            ### 🚀 Go-to-Market Strategy
            
            **Target Customers:**
            1. **Regional Rural Banks** (100+ in India)
            2. **Cooperative Banks** (1,500+)
            3. **Small NBFCs** (10,000+)
            4. **Microfinance Institutions**
            5. **Urban Cooperative Banks**
            
            **Marketing Channels:**
            
            **1. Direct Outreach**
            - LinkedIn targeting bank GMs, CFOs
            - Email campaigns to banking associations
            - RBI registered entity list targeting
            
            **2. Content Marketing**
            - Blog: "How we reduced NPA by 30%"
            - YouTube demos
            - LinkedIn case studies
            
            **3. Partnerships**
            - Banking software vendors (Finacle, Flexcube)
            - Banking consultants
            - Audit firms (Big 4)
            
            **4. Events**
            - Banking technology conferences
            - RBI workshops
            - IBA (Indian Banks Association) events
            
            **Launch Plan:**
            
            **Month 1-2: Beta**
            - 10 pilot banks (free)
            - Collect testimonials
            - Refine product
            
            **Month 3-4: Launch**
            - Press release
            - LinkedIn campaign
            - 50 demo calls
            
            **Month 5-6: Scale**
            - Hire 2 sales reps
            - Partner with 3 consultants
            - Target 20 paying customers
            
            **Success Metrics:**
            - 100 signups in 6 months
            - 30 paying customers
            - ₹25L ARR
            - 4.5+ star reviews
            """)
    
    else:
        st.markdown("""<div style="text-align:center;padding:3rem;">
        <h2>🚀 Welcome to BankVista</h2>
        <p style="font-size:1.2rem;color:#666;margin-top:1rem;">Revolutionary single-file dynamic dashboard</p>
        <p style="color:#999;margin-top:1rem;">Upload your data or try sample to see the magic</p>
        <p style="color:#3498db;margin-top:2rem;font-size:1.1rem;">✨ One Excel file. All branches. Dropdown magic. ✨</p>
        </div>""", unsafe_allow_html=True)
        
        with st.expander("📋 Required Data Format"):
            st.markdown("""| Column | Description | Example |
|--------|-------------|---------|
| Branch_ID | Unique branch code | B1001 |
| Branch_Name | Full branch name | Mansoorabad |
| Zone | Geographic zone | Telangana |
| Total_Deposits | Deposits in Cr | 110.97 |
| Deposit_Target | Target in Cr | 105.46 |
| Advances | Loans in Cr | 232.29 |
| Advance_Target | Target in Cr | 218.16 |
| NPA_Percent | NPA percentage | 2.8 |
| Profit_Per_Staff | Profit per staff (Lakhs) | 5.44 |
| CASA_Percent | CASA ratio % | 42.3 |
| CD_Ratio | Credit-Deposit % | 72.5 |
| Business_Per_Staff | Business per staff (Cr) | 85.2 |
| Staff_Count | Total staff | 25 |""")

if __name__ == "__main__":
    main()
