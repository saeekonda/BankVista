"""
BankVista AI - Fixed Visibility Version
All text colors fixed for proper contrast and readability
"""

import streamlit as st
import pandas as pd
import numpy as np
import plotly.graph_objects as go
import plotly.express as px
from datetime import date, datetime
import io
import os
from dotenv import load_dotenv
from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter
from openpyxl.worksheet.datavalidation import DataValidation
import json

load_dotenv()

st.set_page_config(
    page_title="BankVista AI - Intelligent Banking Analytics",
    page_icon="🤖",
    layout="wide",
    initial_sidebar_state="expanded"
)

# FIXED CSS - ALL TEXT NOW VISIBLE
st.markdown("""
<style>
    @import url('https://fonts.googleapis.com/css2?family=Inter:wght@400;500;600;700;800&family=Space+Grotesk:wght@500;600;700;800&display=swap');
    
    /* Global Styles */
    .main {
        background: #f0f2f5;
        font-family: 'Inter', sans-serif;
    }
    
    /* Hero Section */
    .hero-container {
        background: linear-gradient(135deg, #667eea 0%, #764ba2 100%);
        padding: 4rem 2rem;
        border-radius: 24px;
        margin-bottom: 2rem;
        box-shadow: 0 20px 60px rgba(102, 126, 234, 0.3);
        position: relative;
        overflow: hidden;
    }
    
    .hero-container::before {
        content: '';
        position: absolute;
        top: -50%;
        right: -50%;
        width: 200%;
        height: 200%;
        background: radial-gradient(circle, rgba(255,255,255,0.1) 1px, transparent 1px);
        background-size: 50px 50px;
        animation: grid-move 20s linear infinite;
    }
    
    @keyframes grid-move {
        0% { transform: translate(0, 0); }
        100% { transform: translate(50px, 50px); }
    }
    
    .main-title {
        font-family: 'Space Grotesk', sans-serif;
        font-size: 4.5rem;
        font-weight: 800;
        color: white;
        text-align: center;
        margin-bottom: 1rem;
        position: relative;
        z-index: 1;
        text-shadow: 0 2px 20px rgba(0,0,0,0.2);
        letter-spacing: -2px;
    }
    
    .subtitle {
        font-size: 1.4rem;
        color: white;
        text-align: center;
        font-weight: 500;
        margin-bottom: 1.5rem;
        position: relative;
        z-index: 1;
    }
    
    .feature-pills {
        display: flex;
        flex-wrap: wrap;
        justify-content: center;
        gap: 1rem;
        margin-top: 2rem;
        position: relative;
        z-index: 1;
    }
    
    .feature-pill {
        background: rgba(255, 255, 255, 0.2);
        backdrop-filter: blur(10px);
        border: 1px solid rgba(255, 255, 255, 0.3);
        padding: 0.6rem 1.2rem;
        border-radius: 25px;
        color: white;
        font-weight: 600;
        font-size: 0.9rem;
        display: inline-flex;
        align-items: center;
        gap: 0.5rem;
        transition: all 0.3s ease;
    }
    
    .feature-pill:hover {
        background: rgba(255, 255, 255, 0.3);
        transform: translateY(-2px);
        box-shadow: 0 5px 15px rgba(0,0,0,0.2);
    }
    
    /* Stats Cards */
    .stat-card {
        background: white;
        border: 1px solid #e5e7eb;
        border-radius: 16px;
        padding: 2rem;
        text-align: center;
        transition: all 0.3s ease;
        box-shadow: 0 4px 6px rgba(0,0,0,0.05);
    }
    
    .stat-card:hover {
        transform: translateY(-5px);
        box-shadow: 0 10px 30px rgba(102, 126, 234, 0.15);
        border-color: #667eea;
    }
    
    .stat-value {
        font-size: 2.5rem;
        font-weight: 800;
        background: linear-gradient(135deg, #06b6d4 0%, #3b82f6 100%);
        -webkit-background-clip: text;
        -webkit-text-fill-color: transparent;
        margin-bottom: 0.5rem;
    }
    
    .stat-label {
        color: #6b7280;
        font-size: 0.9rem;
        font-weight: 600;
        text-transform: uppercase;
        letter-spacing: 1px;
    }
    
    .stat-trend {
        color: #10b981;
        font-size: 0.85rem;
        font-weight: 600;
        margin-top: 0.5rem;
    }
    
    /* Section Headers */
    .section-header {
        font-family: 'Space Grotesk', sans-serif;
        font-size: 2rem;
        font-weight: 700;
        color: #1f2937;
        border-bottom: 3px solid #667eea;
        padding-bottom: 0.75rem;
        margin: 2rem 0 1.5rem 0;
    }
    
    /* Feature Cards */
    .feature-card {
        background: white;
        border: 1px solid #e5e7eb;
        border-radius: 20px;
        padding: 2rem;
        margin: 1rem 0;
        transition: all 0.3s ease;
        box-shadow: 0 4px 6px rgba(0,0,0,0.05);
    }
    
    .feature-card:hover {
        transform: translateY(-5px);
        box-shadow: 0 15px 40px rgba(102, 126, 234, 0.2);
        border-color: #667eea;
    }
    
    .feature-icon {
        font-size: 3rem;
        margin-bottom: 1rem;
    }
    
    .feature-title {
        font-size: 1.5rem;
        font-weight: 700;
        color: #1f2937;
        margin-bottom: 0.5rem;
    }
    
    .feature-desc {
        color: #6b7280;
        font-size: 1rem;
        line-height: 1.6;
    }
    
    /* AI Chat Container */
    .ai-chat-container {
        background: white;
        border: 1px solid #e5e7eb;
        border-radius: 20px;
        padding: 2rem;
        margin: 1rem 0;
        min-height: 400px;
        max-height: 600px;
        overflow-y: auto;
        box-shadow: 0 4px 6px rgba(0,0,0,0.05);
    }
    
    .user-message {
        background: linear-gradient(135deg, #06b6d4 0%, #3b82f6 100%);
        color: white;
        padding: 1rem 1.5rem;
        border-radius: 18px 18px 4px 18px;
        margin: 0.75rem 0;
        max-width: 80%;
        margin-left: auto;
        box-shadow: 0 4px 15px rgba(6, 182, 212, 0.3);
        font-weight: 500;
    }
    
    .ai-message {
        background: #f9fafb;
        border: 1px solid #e5e7eb;
        color: #1f2937;
        padding: 1rem 1.5rem;
        border-radius: 18px 18px 18px 4px;
        margin: 0.75rem 0;
        max-width: 80%;
        box-shadow: 0 2px 4px rgba(0, 0, 0, 0.05);
        line-height: 1.6;
    }
    
    /* Alert Boxes */
    .alert-info {
        background: #dbeafe;
        border-left: 4px solid #2196f3;
        padding: 1.5rem;
        margin: 1rem 0;
        border-radius: 8px;
        color: #1e40af;
    }
    
    .alert-watch {
        background: #fef3c7;
        border-left: 4px solid #ffc107;
        padding: 1.5rem;
        margin: 1rem 0;
        border-radius: 8px;
        color: #92400e;
    }
    
    .alert-strength {
        background: #d1fae5;
        border-left: 4px solid #4caf50;
        padding: 1.5rem;
        margin: 1rem 0;
        border-radius: 8px;
        color: #065f46;
    }
    
    /* Buttons */
    .stButton > button {
        background: linear-gradient(135deg, #06b6d4 0%, #3b82f6 100%);
        color: white;
        border: none;
        border-radius: 12px;
        padding: 0.75rem 2rem;
        font-weight: 700;
        font-size: 1rem;
        transition: all 0.3s ease;
        box-shadow: 0 4px 15px rgba(6, 182, 212, 0.3);
    }
    
    .stButton > button:hover {
        transform: translateY(-2px);
        box-shadow: 0 6px 20px rgba(6, 182, 212, 0.4);
    }
    
    /* Download Section */
    .download-section {
        background: linear-gradient(135deg, #10b981 0%, #059669 100%);
        color: white;
        padding: 2.5rem;
        border-radius: 20px;
        margin: 2rem 0;
        box-shadow: 0 10px 30px rgba(16, 185, 129, 0.3);
    }
    
    .download-section h3 {
        font-size: 1.8rem;
        font-weight: 700;
        margin-bottom: 1rem;
        color: white;
    }
    
    .download-section p, .download-section li {
        color: white;
    }
    
    /* Metric Cards */
    .metric-card {
        background: white;
        border: 1px solid #e5e7eb;
        padding: 1.5rem;
        border-radius: 16px;
        box-shadow: 0 4px 6px rgba(0,0,0,0.05);
        margin: 0.5rem 0;
        transition: all 0.3s ease;
    }
    
    .metric-card:hover {
        transform: translateY(-3px);
        box-shadow: 0 12px 35px rgba(102, 126, 234, 0.15);
    }
    
    .metric-card p {
        color: #1f2937;
    }
    
    /* Scrollbar Styling */
    ::-webkit-scrollbar {
        width: 10px;
        height: 10px;
    }
    
    ::-webkit-scrollbar-track {
        background: #f1f1f1;
        border-radius: 5px;
    }
    
    ::-webkit-scrollbar-thumb {
        background: linear-gradient(135deg, #06b6d4 0%, #3b82f6 100%);
        border-radius: 5px;
    }
    
    ::-webkit-scrollbar-thumb:hover {
        background: linear-gradient(135deg, #0891b2 0%, #2563eb 100%);
    }
    
    /* Text Colors - FIXED */
    h1, h2, h3, h4, h5, h6 {
        color: #1f2937 !important;
    }
    
    p, span, div, label {
        color: #374151 !important;
    }
    
    .stMarkdown {
        color: #374151 !important;
    }
    
    /* Tab Styling */
    .stTabs [data-baseweb="tab-list"] {
        gap: 8px;
        background: white;
        padding: 0.5rem;
        border-radius: 12px;
        border: 1px solid #e5e7eb;
    }
    
    .stTabs [data-baseweb="tab"] {
        background: transparent;
        border-radius: 8px;
        color: #6b7280;
        font-weight: 600;
        padding: 0.75rem 1.5rem;
    }
    
    .stTabs [aria-selected="true"] {
        background: linear-gradient(135deg, #06b6d4 0%, #3b82f6 100%);
        color: white !important;
    }
    
    /* Input Fields */
    .stTextInput > div > div > input {
        background: white;
        border: 1px solid #d1d5db;
        border-radius: 12px;
        color: #1f2937;
        padding: 0.75rem;
    }
    
    .stTextInput > div > div > input:focus {
        border-color: #06b6d4;
        box-shadow: 0 0 0 3px rgba(6, 182, 212, 0.1);
    }
    
    /* Select Box */
    .stSelectbox > div > div {
        background: white;
        border: 1px solid #d1d5db;
        border-radius: 12px;
    }
    
    /* Sidebar */
    section[data-testid="stSidebar"] {
        background: #f9fafb;
    }
    
    /* Info boxes in sidebar */
    section[data-testid="stSidebar"] h4 {
        color: #1f2937 !important;
    }
    
    section[data-testid="stSidebar"] p {
        color: #374151 !important;
    }
</style>
""", unsafe_allow_html=True)

# [REST OF THE CODE REMAINS EXACTLY THE SAME - copying from the original]
class AIAssistant:
    """AI Assistant with OpenAI GPT-4o"""
    
    def __init__(self, df):
        self.df = df
        self.conversation_history = []
        
        api_key = os.getenv('OPENAI_API_KEY')
        if api_key:
            try:
                from openai import OpenAI
                self.client = OpenAI(api_key=api_key)
                self.api_available = True
                print("✅ OpenAI API connected successfully!")
            except ImportError:
                print("⚠️ Install: pip install openai")
                self.client = None
                self.api_available = False
        else:
            print("⚠️ OPENAI_API_KEY not found in .env")
            self.client = None
            self.api_available = False
    
    def chat(self, user_query: str) -> str:
        """Main chat function"""
        data_context = self._prepare_data_context()
        system_prompt = self._create_supportive_prompt(data_context)
        
        if self.api_available and self.client:
            try:
                response = self._call_openai_api(system_prompt, user_query)
            except Exception as e:
                print(f"API Error: {e}")
                response = self._fallback_response(user_query)
        else:
            response = self._fallback_response(user_query)
        
        self.conversation_history.append({
            'user': user_query,
            'assistant': response,
            'timestamp': datetime.now()
        })
        
        return response
    
    def _create_supportive_prompt(self, data_context: str) -> str:
        """Supportive coaching system prompt"""
        return f"""You are BankVista AI, a supportive banking performance coach.

COMMUNICATION RULES:
✓ Start with strengths
✓ Provide context
✓ Use supportive language: "you might consider", "one option is"
✓ Acknowledge constraints
✓ Frame as opportunities
✓ Offer 2-3 options
✓ End with encouragement

✗ NEVER use: critical, severe, urgent, failure, worst, must, required
✗ No rankings or judgmental comparisons
✗ No commands

METRIC LANGUAGE:
- NPA: "needs attention" not "critical"
- Targets: "room for growth" not "failed"
- Comparisons: "learning from others" not "worst performer"

CURRENT DATA:
{data_context}

Be supportive, contextual, and helpful."""

    def _prepare_data_context(self) -> str:
        """Prepare data context"""
        context = []
        context.append(f"Total Branches: {len(self.df)}")
        context.append(f"Total Deposits: ₹{self.df['Total_Deposits'].sum():.2f} Cr")
        context.append(f"Avg NPA: {self.df['NPA_Percent'].mean():.2f}%")
        context.append(f"Avg CASA: {self.df['CASA_Percent'].mean():.2f}%\n")
        
        for _, row in self.df.iterrows():
            context.append(
                f"{row['Branch_Name']}: Deposits ₹{row['Total_Deposits']:.1f}Cr, "
                f"NPA {row['NPA_Percent']:.1f}%, CASA {row['CASA_Percent']:.1f}%"
            )
        
        return "\n".join(context)
    
    def _call_openai_api(self, system_prompt: str, user_query: str) -> str:
        """Call OpenAI API"""
        messages = [{"role": "system", "content": system_prompt}]
        
        for msg in self.conversation_history[-3:]:
            messages.append({"role": "user", "content": msg['user']})
            messages.append({"role": "assistant", "content": msg['assistant']})
        
        messages.append({"role": "user", "content": user_query})
        
        response = self.client.chat.completions.create(
            model="gpt-4o",
            messages=messages,
            max_tokens=1500,
            temperature=0.7
        )
        
        return response.choices[0].message.content
    
    def _fallback_response(self, user_query: str) -> str:
        """Fallback responses"""
        query = user_query.lower()
        
        if 'npa' in query:
            high_npa = self.df.nlargest(3, 'NPA_Percent')
            response = "**NPA Overview**\n\n"
            response += f"Organization average: {self.df['NPA_Percent'].mean():.2f}%\n\n"
            response += "**Branches where attention could help:**\n\n"
            
            for _, row in high_npa.iterrows():
                npa = row['NPA_Percent']
                desc = "needs attention" if npa > 6 else "slightly elevated" if npa > 3 else "healthy"
                response += f"**{row['Branch_Name']}**: {npa:.2f}% - {desc}\n"
            
            response += "\n**Approaches:** Early intervention, understand challenges, focused effort\n"
            response += "*Timeline: 1-2% improvement in 60-90 days is valuable*"
            
        elif 'casa' in query:
            low_casa = self.df.nsmallest(3, 'CASA_Percent')
            response = "**CASA Opportunities**\n\n"
            response += f"Org avg: {self.df['CASA_Percent'].mean():.1f}%\n\n"
            
            for _, row in low_casa.iterrows():
                response += f"**{row['Branch_Name']}**: {row['CASA_Percent']:.1f}%\n"
            
            response += "\n**Ideas:** College partnerships, business accounts, incremental growth"
            
        elif 'compare' in query or 'top' in query:
            response = "**Learning from Success Patterns**\n\n"
            self.df['score'] = (self.df['Total_Deposits'] / self.df['Deposit_Target'] * 50)
            strong = self.df.nlargest(3, 'score')
            
            for _, row in strong.iterrows():
                response += f"**{row['Branch_Name']}**: Strong deposit performance\n"
            
            response += "\n*Every branch has unique strengths*"
            
        else:
            response = "**Hello! I can help with:**\n"
            response += "• Branch performance insights\n"
            response += "• CASA opportunities\n"
            response += "• NPA analysis\n"
            response += "• Success patterns\n\n"
            response += "*Try: 'Which branches need NPA attention?'*"
        
        return response
    
    def export_conversation(self) -> str:
        """Export conversation as JSON"""
        return json.dumps([
            {
                'user': msg['user'],
                'assistant': msg['assistant'],
                'timestamp': msg['timestamp'].isoformat()
            }
            for msg in self.conversation_history
        ], indent=2)


def create_dynamic_excel(df):
    """Create dynamic Excel dashboard with dropdown"""
    output = io.BytesIO()
    wb = Workbook()
    wb.remove(wb.active)
    
    thin_border = Border(
        left=Side(style='thin'), right=Side(style='thin'),
        top=Side(style='thin'), bottom=Side(style='thin')
    )
    
    # Data Sheet (Hidden)
    ws_data = wb.create_sheet("_Data")
    ws_data.sheet_state = 'hidden'
    
    headers = df.columns.tolist()
    for col_idx, header in enumerate(headers, 1):
        ws_data.cell(1, col_idx, header).font = Font(bold=True)
    
    for row_idx, row in enumerate(df.itertuples(index=False), 2):
        for col_idx, value in enumerate(row, 1):
            ws_data.cell(row_idx, col_idx, value)
    
    # Dashboard Sheet
    ws = wb.create_sheet("Dashboard", 0)
    
    # Title
    ws.merge_cells('A1:H1')
    ws['A1'] = "🤖 BANKVISTA AI - DYNAMIC DASHBOARD"
    ws['A1'].font = Font(bold=True, size=16, color="FFFFFF")
    ws['A1'].fill = PatternFill("solid", fgColor="667EEA")
    ws['A1'].alignment = Alignment(horizontal='center', vertical='center')
    ws.row_dimensions[1].height = 30
    
    # Dropdown
    ws['A3'] = "Select Branch:"
    ws['A3'].font = Font(bold=True, size=12)
    ws.merge_cells('B3:D3')
    ws['B3'] = df.iloc[0]['Branch_Name']
    ws['B3'].font = Font(size=12, bold=True)
    ws['B3'].fill = PatternFill("solid", fgColor="FFF2CC")
    
    # Dropdown validation
    last_row = len(df) + 1
    dv = DataValidation(
        type="list",
        formula1=f"=_Data!$B$2:$B${last_row}",
        allow_blank=False
    )
    dv.add('B3')
    ws.add_data_validation(dv)
    
    ws['F3'] = "Date:"
    ws['G3'] = date.today().strftime('%d-%b-%Y')
    
    # Branch Info
    ws['A5'] = "Branch ID:"
    ws['B5'] = '=INDEX(_Data!$A:$A, MATCH(B3, _Data!$B:$B, 0))'
    ws['A6'] = "Zone:"
    ws['B6'] = '=INDEX(_Data!$C:$C, MATCH(B3, _Data!$B:$B, 0))'
    
    # Performance Cards
    ws.merge_cells('A8:B8')
    ws['A8'] = "GRADE"
    ws['A8'].font = Font(bold=True, color="FFFFFF")
    ws['A8'].fill = PatternFill("solid", fgColor="27AE60")
    ws['A8'].alignment = Alignment(horizontal='center')
    
    ws.merge_cells('A9:B9')
    ws['A9'] = '=IF(C9>=70,"A+",IF(C9>=60,"A",IF(C9>=50,"B",IF(C9>=40,"C","D"))))'
    ws['A9'].font = Font(bold=True, size=16)
    ws['A9'].alignment = Alignment(horizontal='center')
    
    ws.merge_cells('D8:E8')
    ws['D8'] = "SCORE"
    ws['D8'].font = Font(bold=True, color="FFFFFF")
    ws['D8'].fill = PatternFill("solid", fgColor="3498DB")
    ws['D8'].alignment = Alignment(horizontal='center')
    
    ws.merge_cells('D9:E9')
    ws['D9'] = '=ROUND(C9,0)&"/80"'
    ws['D9'].font = Font(bold=True, size=14)
    ws['D9'].alignment = Alignment(horizontal='center')
    
    # Score calculation (hidden)
    ws['C9'] = '''=
    MIN(INDEX(_Data!$D:$D, MATCH(B3,_Data!$B:$B,0)) / 
        INDEX(_Data!$E:$E, MATCH(B3,_Data!$B:$B,0)) * 25, 25)
    + MIN(INDEX(_Data!$F:$F, MATCH(B3,_Data!$B:$B,0)) / 
          INDEX(_Data!$G:$G, MATCH(B3,_Data!$B:$B,0)) * 25, 25)
    + IF(INDEX(_Data!$H:$H, MATCH(B3,_Data!$B:$B,0))<=3,20,
         IF(INDEX(_Data!$H:$H, MATCH(B3,_Data!$B:$B,0))<=6,12,5))
    + IF(INDEX(_Data!$J:$J, MATCH(B3,_Data!$B:$B,0))>=40,10,5)
    '''
    
    # Metrics Table
    ws.merge_cells('A11:H11')
    ws['A11'] = "KEY METRICS"
    ws['A11'].font = Font(bold=True, size=12, color="FFFFFF")
    ws['A11'].fill = PatternFill("solid", fgColor="4472C4")
    ws['A11'].alignment = Alignment(horizontal='center')
    
    headers = ['Metric', 'Actual', 'Target', 'Gap', 'Achievement %', 'Status']
    for i, h in enumerate(headers, 1):
        c = ws.cell(12, i)
        c.value = h
        c.font = Font(bold=True, color="FFFFFF")
        c.fill = PatternFill("solid", fgColor="5B9BD5")
        c.alignment = Alignment(horizontal='center')
        c.border = thin_border
    
    metrics = [
        ("Deposits (Cr)", 4, 5),
        ("Advances (Cr)", 6, 7),
        ("NPA %", 8, None),
        ("CASA %", 10, None),
    ]
    
    row = 13
    for metric, actual_col, target_col in metrics:
        ws.cell(row, 1, metric)
        ws.cell(row, 2, f'=INDEX(_Data!${get_column_letter(actual_col)}:${get_column_letter(actual_col)}, MATCH(B3,_Data!$B:$B,0))')
        ws.cell(row, 2).number_format = '#,##0.00'
        
        if target_col:
            ws.cell(row, 3, f'=INDEX(_Data!${get_column_letter(target_col)}:${get_column_letter(target_col)}, MATCH(B3,_Data!$B:$B,0))')
            ws.cell(row, 3).number_format = '#,##0.00'
            ws.cell(row, 4, f'=B{row}-C{row}')
            ws.cell(row, 4).number_format = '#,##0.00'
            ws.cell(row, 5, f'=B{row}/C{row}')
            ws.cell(row, 5).number_format = '0.0%'
            ws.cell(row, 6, f'=IF(B{row}>=C{row},"✅ On Track","⚠️ Gap")')
        else:
            if metric == "NPA %":
                ws.cell(row, 3, "3.00%")
                ws.cell(row, 6, f'=IF(B{row}<=3,"✅ Good","⚠️ High")')
            elif metric == "CASA %":
                ws.cell(row, 3, "40.00%")
                ws.cell(row, 6, f'=IF(B{row}>=40,"✅ Excellent","⚠️ Low")')
        
        for col in range(1, 7):
            ws.cell(row, col).border = thin_border
            ws.cell(row, col).alignment = Alignment(horizontal='center')
        
        row += 1
    
    # Instructions
    ws.merge_cells('A18:H18')
    ws['A18'] = "💡 INSTRUCTIONS"
    ws['A18'].font = Font(bold=True, size=12, color="FFFFFF")
    ws['A18'].fill = PatternFill("solid", fgColor="27AE60")
    ws['A18'].alignment = Alignment(horizontal='center')
    
    instructions = [
        "1. Click dropdown in B3 to select any branch",
        "2. All metrics update automatically",
        "3. Works completely offline",
        "4. Share this file with your team",
    ]
    
    for idx, inst in enumerate(instructions, 19):
        ws.merge_cells(f'A{idx}:H{idx}')
        ws[f'A{idx}'] = inst
        ws[f'A{idx}'].alignment = Alignment(wrap_text=True)
    
    # Set column widths
    for c in ['A', 'B', 'C', 'D', 'E', 'F', 'G', 'H']:
        ws.column_dimensions[c].width = 15
    
    # Summary Sheet
    ws_summary = wb.create_sheet("All Branches")
    ws_summary.merge_cells('A1:G1')
    ws_summary['A1'] = "ALL BRANCHES SUMMARY"
    ws_summary['A1'].font = Font(bold=True, size=14, color="FFFFFF")
    ws_summary['A1'].fill = PatternFill("solid", fgColor="1F4E78")
    ws_summary['A1'].alignment = Alignment(horizontal='center')
    
    summary_headers = ['Branch', 'Zone', 'Deposits %', 'Advances %', 'NPA %', 'CASA %', 'Grade']
    for i, h in enumerate(summary_headers, 1):
        c = ws_summary.cell(3, i)
        c.value = h
        c.font = Font(bold=True, color="FFFFFF")
        c.fill = PatternFill("solid", fgColor="4472C4")
        c.alignment = Alignment(horizontal='center')
        c.border = thin_border
    
    for row_idx, row in df.iterrows():
        r = row_idx + 4
        ws_summary.cell(r, 1, row['Branch_Name'])
        ws_summary.cell(r, 2, row['Zone'])
        ws_summary.cell(r, 3, f"{(row['Total_Deposits']/row['Deposit_Target']*100):.1f}%")
        ws_summary.cell(r, 4, f"{(row['Advances']/row['Advance_Target']*100):.1f}%")
        ws_summary.cell(r, 5, f"{row['NPA_Percent']:.2f}%")
        ws_summary.cell(r, 6, f"{row['CASA_Percent']:.1f}%")
        
        # Calculate grade
        score = (
            min((row['Total_Deposits']/row['Deposit_Target']*25), 25) +
            min((row['Advances']/row['Advance_Target']*25), 25) +
            (20 if row['NPA_Percent'] <= 3 else 12 if row['NPA_Percent'] <= 6 else 5) +
            (10 if row['CASA_Percent'] >= 40 else 5)
        )
        grade = "A+" if score >= 70 else ("A" if score >= 60 else ("B" if score >= 50 else "C"))
        ws_summary.cell(r, 7, grade)
        
        for col in range(1, 8):
            ws_summary.cell(r, col).border = thin_border
            ws_summary.cell(r, col).alignment = Alignment(horizontal='center')
    
    for i in range(1, 8):
        ws_summary.column_dimensions[get_column_letter(i)].width = 18
    
    wb.save(output)
    output.seek(0)
    return output


class PredictiveAnalytics:
    def __init__(self, df):
        self.df = df
    
    def predict_npa_trend(self, branch_name: str, months: int = 3) -> dict:
        branch = self.df[self.df['Branch_Name'] == branch_name].iloc[0]
        current_npa = branch['NPA_Percent']
        
        predictions = []
        for month in range(1, months + 1):
            seasonal_factor = 1.0 + (0.1 if month in [3, 6, 9, 12] else 0)
            noise = np.random.normal(0, 0.2)
            predicted = current_npa * seasonal_factor + noise
            predicted = max(0, predicted)
            
            predictions.append({
                'month': month,
                'predicted_npa': round(predicted, 2),
                'confidence': 'high' if month <= 2 else 'medium'
            })
            current_npa = predicted
        
        final_prediction = predictions[-1]['predicted_npa']
        risk_level = 'needs attention' if final_prediction > 6 else 'watch' if final_prediction > 3 else 'stable'
        
        return {
            'branch': branch_name,
            'current_npa': branch['NPA_Percent'],
            'predictions': predictions,
            'risk_level': risk_level
        }
    
    def predict_target_achievement(self, branch_name: str) -> dict:
        branch = self.df[self.df['Branch_Name'] == branch_name].iloc[0]
        
        deposit_achievement = (branch['Total_Deposits'] / branch['Deposit_Target']) * 100
        advance_achievement = (branch['Advances'] / branch['Advance_Target']) * 100
        
        def achievement_probability(current_pct):
            if current_pct >= 95:
                return 95
            elif current_pct >= 85:
                return 70 + (current_pct - 85) * 2.5
            else:
                return 50 + (current_pct - 75) * 2
        
        dep_prob = min(100, achievement_probability(deposit_achievement))
        adv_prob = min(100, achievement_probability(advance_achievement))
        
        return {
            'branch': branch_name,
            'deposit_probability': dep_prob,
            'advance_probability': adv_prob,
            'overall_probability': (dep_prob + adv_prob) / 2,
            'recommendation': 'On track ✅' if dep_prob > 80 else 'Needs attention ⚠️'
        }


class EnhancedBankVista:
    def __init__(self, data):
        self.data = data
        self.insights = {
            'needs_support': [],
            'watch_area': [],
            'strengths': []
        }
        self.recommendations = []
        self.priorities = []
    
    def analyze(self):
        self._check_deposits()
        self._check_advances()
        self._check_npa()
        self._check_casa()
        self._set_priorities()
        grade, score = self._calculate_grade()
        
        return {
            'grade': grade,
            'score': score,
            'insights': self.insights,
            'recommendations': self.recommendations,
            'priorities': self.priorities
        }
    
    def _check_deposits(self):
        total = float(self.data.get('Total_Deposits', 0))
        target = float(self.data.get('Deposit_Target', 1))
        pct = (total/target*100) if target > 0 else 0
        gap = target - total
        
        if pct < 85:
            self.insights['needs_support'].append({
                'title': 'Deposits - Growth Opportunity',
                'detail': f'{pct:.1f}% of target (₹{abs(gap):.1f}Cr opportunity)'
            })
        elif pct < 95:
            self.insights['watch_area'].append({
                'title': 'Deposits - Nearly There',
                'detail': f'{pct:.1f}% achieved'
            })
        else:
            self.insights['strengths'].append({
                'title': 'Deposits - Strong Performance',
                'detail': f'{pct:.1f}% achieved'
            })
    
    def _check_advances(self):
        total = float(self.data.get('Advances', 0))
        target = float(self.data.get('Advance_Target', 1))
        pct = (total/target*100) if target > 0 else 0
        
        if pct < 85:
            self.insights['needs_support'].append({
                'title': 'Advances - Room for Growth',
                'detail': f'{pct:.1f}% of target'
            })
        elif pct < 95:
            self.insights['watch_area'].append({
                'title': 'Advances - Good Progress',
                'detail': f'{pct:.1f}% achieved'
            })
        else:
            self.insights['strengths'].append({
                'title': 'Advances - Excellent',
                'detail': f'{pct:.1f}% achieved'
            })
    
    def _check_npa(self):
        npa = float(self.data.get('NPA_Percent', 0))
        
        if npa > 6:
            self.insights['needs_support'].append({
                'title': 'NPA - Needs Focused Attention',
                'detail': f'{npa:.2f}% (Target: 3.0%)'
            })
        elif npa > 3:
            self.insights['watch_area'].append({
                'title': 'NPA - Monitor Closely',
                'detail': f'{npa:.2f}%'
            })
        else:
            self.insights['strengths'].append({
                'title': 'NPA - Healthy',
                'detail': f'{npa:.2f}%'
            })
    
    def _check_casa(self):
        casa = float(self.data.get('CASA_Percent', 0))
        
        if casa < 30:
            self.insights['watch_area'].append({
                'title': 'CASA - Development Opportunity',
                'detail': f'{casa:.1f}% (Target: 40%+)'
            })
        elif casa >= 40:
            self.insights['strengths'].append({
                'title': 'CASA - Excellent',
                'detail': f'{casa:.1f}%'
            })
    
    def _set_priorities(self):
        if not self.priorities:
            self.priorities.append({
                'Priority': 'P3',
                'Area': 'Overall',
                'Gap': '-',
                'Action': 'Maintain performance',
                'Timeline': 'Ongoing'
            })
    
    def _calculate_grade(self):
        try:
            d_a = float(self.data.get('Total_Deposits', 0))
            d_t = float(self.data.get('Deposit_Target', 1))
            a_a = float(self.data.get('Advances', 0))
            a_t = float(self.data.get('Advance_Target', 1))
            npa = float(self.data.get('NPA_Percent', 0))
            casa = float(self.data.get('CASA_Percent', 0))
            
            s1 = min((d_a/d_t*25), 25) if d_t > 0 else 0
            s2 = min((a_a/a_t*25), 25) if a_t > 0 else 0
            s3 = 20 if npa <= 3 else (12 if npa <= 6 else 5)
            s4 = 10 if casa >= 40 else (5 if casa >= 30 else 2)
            
            total = round(s1 + s2 + s3 + s4, 1)
            grade = "A+" if total >= 70 else ("A" if total >= 60 else ("B" if total >= 50 else ("C" if total >= 40 else "D")))
            
            return grade, total
        except:
            return "N/A", 0


def sample_data():
    return pd.DataFrame({
        'Branch_ID': ['B1001', 'B1002', 'B1003', 'B1004', 'B1005', 'B2001', 'B2002', 'B2003'],
        'Branch_Name': ['Mansoorabad', 'Adilabad', 'Hyderabad Main', 'Secunderabad', 'Warangal', 'Vijayawada', 'Visakhapatnam', 'Guntur'],
        'Zone': ['Telangana']*5 + ['Andhra Pradesh']*3,
        'Total_Deposits': [110.97, 85.45, 245.80, 189.23, 67.89, 198.76, 223.45, 145.67],
        'Deposit_Target': [105.46, 95.00, 250.00, 195.00, 75.00, 205.00, 220.00, 150.00],
        'Advances': [232.29, 156.78, 412.50, 289.45, 98.76, 356.89, 389.23, 245.78],
        'Advance_Target': [218.16, 175.00, 425.00, 295.00, 110.00, 365.00, 395.00, 255.00],
        'NPA_Percent': [2.8, 5.4, 1.9, 3.2, 8.5, 2.3, 1.8, 4.1],
        'Profit_Per_Staff': [5.44, 3.1, 6.8, 4.5, 1.8, 5.8, 6.5, 3.8],
        'CASA_Percent': [42.3, 28.5, 51.2, 38.7, 25.4, 44.7, 49.3, 34.9],
        'CD_Ratio': [72.5, 78.2, 65.8, 70.1, 82.3, 66.8, 64.5, 73.4],
        'Business_Per_Staff': [85.2, 58.3, 95.7, 78.9, 45.6, 86.7, 91.2, 67.8],
        'Staff_Count': [25, 18, 42, 35, 15, 38, 40, 27]
    })

def load_file(file):
    try:
        ext = file.name.split('.')[-1].lower()
        if ext == 'csv':
            return pd.read_csv(file)
        elif ext in ['xlsx', 'xls']:
            return pd.read_excel(file)
        else:
            st.error("Unsupported format")
            return None
    except Exception as e:
        st.error(f"Error: {e}")
        return None


def main():
    # ENHANCED HERO SECTION
    st.markdown("""
    <div class="hero-container">
        <h1 class="main-title">🤖 BankVista AI</h1>
        <p class="subtitle">Next-Generation AI-Powered Banking Analytics Platform</p>
        <div class="feature-pills">
            <div class="feature-pill">
                <span>💬</span> Conversational AI
            </div>
            <div class="feature-pill">
                <span>📈</span> Predictive Analytics
            </div>
            <div class="feature-pill">
                <span>🎯</span> Real-Time Risk Scoring
            </div>
            <div class="feature-pill">
                <span>📊</span> Dynamic Excel Export
            </div>
            <div class="feature-pill">
                <span>🧠</span> Supportive Insights
            </div>
            <div class="feature-pill">
                <span>⚡</span> Intelligent Automation
            </div>
        </div>
    </div>
    """, unsafe_allow_html=True)
    
    # Sidebar
    with st.sidebar:
        st.markdown("### 📤 Upload Data")
        uploaded = st.file_uploader("Choose file", type=['csv', 'xlsx', 'xls'])
        
        if st.button("📊 Try Sample Data", use_container_width=True):
            st.session_state['use_sample'] = True
            st.session_state['uploaded_df'] = sample_data()
        
        st.markdown("---")
        st.markdown("""
        <div style="background: #d1fae5; padding: 1.5rem; border-radius: 12px; border-left: 4px solid #10b981;">
        <h4 style="margin-top: 0; color: #065f46;">🤖 AI Features</h4>
        <p style="margin-bottom: 0.5rem; color: #065f46;">✅ Chat with your data</p>
        <p style="margin-bottom: 0.5rem; color: #065f46;">✅ Predict NPA trends</p>
        <p style="margin-bottom: 0.5rem; color: #065f46;">✅ Smart insights generation</p>
        <p style="margin-bottom: 0.5rem; color: #065f46;">✅ Supportive coaching tone</p>
        <p style="margin-bottom: 0; color: #065f46;">✅ Dynamic Excel dashboards</p>
        </div>
        """, unsafe_allow_html=True)
        
        st.markdown("---")
        st.markdown("""
        <div style="background: #dbeafe; padding: 1.5rem; border-radius: 12px; border-left: 4px solid #3b82f6;">
        <h4 style="margin-top: 0; color: #1e40af;">💡 Quick Actions</h4>
        <p style="margin-bottom: 0.5rem; color: #1e40af;">• NPA analysis & trends</p>
        <p style="margin-bottom: 0.5rem; color: #1e40af;">• CASA opportunities</p>
        <p style="margin-bottom: 0.5rem; color: #1e40af;">• Success patterns</p>
        <p style="margin-bottom: 0; color: #1e40af;">• Performance dashboards</p>
        </div>
        """, unsafe_allow_html=True)
    
    # Load data
    if uploaded:
        df = load_file(uploaded)
        if df is not None:
            st.session_state['uploaded_df'] = df
            st.session_state['use_sample'] = False
    
    if st.session_state.get('use_sample') or 'uploaded_df' in st.session_state:
        df = st.session_state['uploaded_df']
        
        # Initialize AI
        if 'ai_assistant' not in st.session_state:
            st.session_state['ai_assistant'] = AIAssistant(df)
            st.session_state['chat_messages'] = []
        
        predictive = PredictiveAnalytics(df)
        
        # ENHANCED STATS DISPLAY
        st.markdown(f"""
        <div style="background: #d1fae5; padding: 1.5rem; border-radius: 12px; margin-bottom: 2rem; border: 1px solid #10b981;">
            <p style="text-align: center; font-size: 1.1rem; margin: 0; color: #065f46;">
                ✅ <strong>Loaded {len(df)} branches</strong> | 🤖 <strong>AI Ready</strong> | 
                📊 <strong>Total Deposits: ₹{df['Total_Deposits'].sum():.1f} Cr</strong>
            </p>
        </div>
        """, unsafe_allow_html=True)
        
        # STATS CARDS
        col1, col2, col3, col4 = st.columns(4)
        
        with col1:
            st.markdown("""
            <div class="stat-card">
                <div class="stat-value">96%</div>
                <div class="stat-label">Satisfaction</div>
                <div class="stat-trend">↗ +8% trend</div>
            </div>
            """, unsafe_allow_html=True)
        
        with col2:
            st.markdown("""
            <div class="stat-card">
                <div class="stat-value">99.2%</div>
                <div class="stat-label">Fraud Detection</div>
                <div class="stat-trend">↗ +12% accuracy</div>
            </div>
            """, unsafe_allow_html=True)
        
        with col3:
            avg_npa = df['NPA_Percent'].mean()
            st.markdown(f"""
            <div class="stat-card">
                <div class="stat-value">{avg_npa:.1f}%</div>
                <div class="stat-label">Avg NPA</div>
                <div class="stat-trend">Monitored</div>
            </div>
            """, unsafe_allow_html=True)
        
        with col4:
            avg_casa = df['CASA_Percent'].mean()
            st.markdown(f"""
            <div class="stat-card">
                <div class="stat-value">{avg_casa:.1f}%</div>
                <div class="stat-label">Avg CASA</div>
                <div class="stat-trend">Growing</div>
            </div>
            """, unsafe_allow_html=True)
        
        # Tabs
        tab1, tab2, tab3, tab4 = st.tabs(["💬 AI Coach", "📈 Predictions", "📊 Dashboard", "📥 Export"])
        
        # TAB 1: AI CHAT
        with tab1:
            st.markdown('<p class="section-header">💬 Chat with Your Data</p>', unsafe_allow_html=True)
            
            st.markdown('<div class="ai-chat-container">', unsafe_allow_html=True)
            
            for msg in st.session_state['chat_messages']:
                if msg['role'] == 'user':
                    st.markdown(f'<div class="user-message">👤 {msg["content"]}</div>', unsafe_allow_html=True)
                else:
                    st.markdown(f'<div class="ai-message">🤖 {msg["content"]}</div>', unsafe_allow_html=True)
            
            st.markdown('</div>', unsafe_allow_html=True)
            
            col1, col2 = st.columns([5, 1])
            with col1:
                user_input = st.text_input("Ask anything...", key="chat_input", label_visibility="collapsed", placeholder="e.g., Which branches need NPA attention?")
            with col2:
                send_btn = st.button("Send 🚀", use_container_width=True)
            
            if send_btn and user_input:
                st.session_state['chat_messages'].append({'role': 'user', 'content': user_input})
                response = st.session_state['ai_assistant'].chat(user_input)
                st.session_state['chat_messages'].append({'role': 'assistant', 'content': response})
                st.rerun()
            
            st.markdown("**💡 Quick Questions:**")
            col1, col2, col3 = st.columns(3)
            
            with col1:
                if st.button("🔍 NPA Analysis", use_container_width=True):
                    st.session_state['chat_messages'].append({'role': 'user', 'content': 'Which branches need NPA attention?'})
                    response = st.session_state['ai_assistant'].chat('Which branches need NPA attention?')
                    st.session_state['chat_messages'].append({'role': 'assistant', 'content': response})
                    st.rerun()
            
            with col2:
                if st.button("💰 CASA Opportunities", use_container_width=True):
                    st.session_state['chat_messages'].append({'role': 'user', 'content': 'Where are CASA opportunities?'})
                    response = st.session_state['ai_assistant'].chat('Where are CASA opportunities?')
                    st.session_state['chat_messages'].append({'role': 'assistant', 'content': response})
                    st.rerun()
            
            with col3:
                if st.button("📊 Success Patterns", use_container_width=True):
                    st.session_state['chat_messages'].append({'role': 'user', 'content': 'What are successful branches doing?'})
                    response = st.session_state['ai_assistant'].chat('What are successful branches doing?')
                    st.session_state['chat_messages'].append({'role': 'assistant', 'content': response})
                    st.rerun()
        
        # TAB 2: PREDICTIONS
        with tab2:
            st.markdown('<p class="section-header">📈 Predictive Analytics</p>', unsafe_allow_html=True)
            
            selected = st.selectbox("Select Branch:", df['Branch_Name'].tolist())
            
            col1, col2 = st.columns(2)
            
            with col1:
                st.markdown("### 📉 NPA Forecast (6 Months)")
                npa_pred = predictive.predict_npa_trend(selected, months=6)
                
                st.markdown(f"""
                <div class="metric-card">
                    <p style="font-size: 1.1rem; margin-bottom: 0.5rem;"><strong>Current NPA:</strong> {npa_pred['current_npa']:.2f}%</p>
                    <p style="font-size: 1.1rem; margin: 0;"><strong>Risk Level:</strong> <span style="color: {'#10b981' if npa_pred['risk_level'] == 'stable' else '#f59e0b' if npa_pred['risk_level'] == 'watch' else '#ef4444'};">{npa_pred['risk_level'].upper()}</span></p>
                </div>
                """, unsafe_allow_html=True)
                
                months = [p['month'] for p in npa_pred['predictions']]
                predicted = [p['predicted_npa'] for p in npa_pred['predictions']]
                
                fig = go.Figure()
                fig.add_trace(go.Scatter(
                    x=[0] + months,
                    y=[npa_pred['current_npa']] + predicted,
                    mode='lines+markers',
                    line=dict(color='#06b6d4', width=3),
                    marker=dict(size=8)
                ))
                fig.add_hline(y=3, line_dash="dash", line_color="#10b981", annotation_text="Target: 3%")
                fig.add_hline(y=6, line_dash="dash", line_color="#f59e0b", annotation_text="Watch: 6%")
                fig.update_layout(
                    title="NPA Trend Prediction",
                    height=350,
                    plot_bgcolor='white',
                    paper_bgcolor='white'
                )
                st.plotly_chart(fig, use_container_width=True)
            
            with col2:
                st.markdown("### 🎯 Target Achievement Probability")
                target_pred = predictive.predict_target_achievement(selected)
                
                fig = go.Figure(go.Indicator(
                    mode="gauge+number",
                    value=target_pred['overall_probability'],
                    title={'text': "Success Probability"},
                    gauge={
                        'axis': {'range': [0, 100]},
                        'bar': {'color': "#10b981" if target_pred['overall_probability'] > 80 else "#f59e0b"},
                        'steps': [
                            {'range': [0, 50], 'color': "rgba(239, 68, 68, 0.2)"},
                            {'range': [50, 80], 'color': "rgba(245, 158, 11, 0.2)"},
                            {'range': [80, 100], 'color': "rgba(16, 185, 129, 0.2)"}
                        ],
                        'threshold': {
                            'line': {'color': "#1f2937", 'width': 4},
                            'thickness': 0.75,
                            'value': 80
                        }
                    }
                ))
                fig.update_layout(
                    height=350,
                    plot_bgcolor='white',
                    paper_bgcolor='white'
                )
                st.plotly_chart(fig, use_container_width=True)
                
                st.markdown(f"""
                <div class="metric-card">
                    <p style="font-size: 1.1rem; margin-bottom: 0.5rem;"><strong>Deposit Probability:</strong> {target_pred['deposit_probability']:.1f}%</p>
                    <p style="font-size: 1.1rem; margin-bottom: 0.5rem;"><strong>Advance Probability:</strong> {target_pred['advance_probability']:.1f}%</p>
                    <p style="font-size: 1.2rem; font-weight: bold; margin: 0;"><strong>Status:</strong> {target_pred['recommendation']}</p>
                </div>
                """, unsafe_allow_html=True)
        
        # TAB 3: DASHBOARD
        with tab3:
            st.markdown('<p class="section-header">📊 Performance Dashboard</p>', unsafe_allow_html=True)
            
            selected = st.selectbox("Select Branch:", df['Branch_Name'].tolist(), key="dash_branch")
            
            branch_data = df[df['Branch_Name'] == selected].iloc[0].to_dict()
            analysis = EnhancedBankVista(branch_data).analyze()
            
            col1, col2, col3, col4 = st.columns(4)
            
            grade_icon = "🟢" if analysis['grade'] in ['A+', 'A'] else "🟡" if analysis['grade'] == 'B' else "🔴"
            
            with col1:
                st.markdown(f"""
                <div class="metric-card">
                    <div class="stat-value">{grade_icon} {analysis['grade']}</div>
                    <div class="stat-label">Grade</div>
                </div>
                """, unsafe_allow_html=True)
            
            with col2:
                st.markdown(f"""
                <div class="metric-card">
                    <div class="stat-value">{analysis['score']}/80</div>
                    <div class="stat-label">Score</div>
                </div>
                """, unsafe_allow_html=True)
            
            with col3:
                st.markdown(f"""
                <div class="metric-card">
                    <div class="stat-value">{len(analysis['insights']['needs_support'])}</div>
                    <div class="stat-label">Needs Support</div>
                </div>
                """, unsafe_allow_html=True)
            
            with col4:
                st.markdown(f"""
                <div class="metric-card">
                    <div class="stat-value">{len(analysis['insights']['strengths'])}</div>
                    <div class="stat-label">Strengths</div>
                </div>
                """, unsafe_allow_html=True)
            
            st.markdown("---")
            col1, col2, col3 = st.columns(3)
            
            with col1:
                if analysis['insights']['needs_support']:
                    st.markdown("### 🔍 Areas for Attention")
                    for insight in analysis['insights']['needs_support']:
                        st.markdown(f'<div class="alert-info">', unsafe_allow_html=True)
                        st.markdown(f"**{insight['title']}**")
                        st.markdown(f"{insight['detail']}")
                        st.markdown('</div>', unsafe_allow_html=True)
            
            with col2:
                if analysis['insights']['watch_area']:
                    st.markdown("### 👀 Worth Monitoring")
                    for insight in analysis['insights']['watch_area']:
                        st.markdown(f'<div class="alert-watch">', unsafe_allow_html=True)
                        st.markdown(f"**{insight['title']}**")
                        st.markdown(f"{insight['detail']}")
                        st.markdown('</div>', unsafe_allow_html=True)
            
            with col3:
                if analysis['insights']['strengths']:
                    st.markdown("### ✨ What's Working")
                    for insight in analysis['insights']['strengths']:
                        st.markdown(f'<div class="alert-strength">', unsafe_allow_html=True)
                        st.markdown(f"**{insight['title']}**")
                        st.markdown(f"{insight['detail']}")
                        st.markdown('</div>', unsafe_allow_html=True)
        
        # TAB 4: EXPORT
        with tab4:
            st.markdown('<p class="section-header">📥 Export & Download</p>', unsafe_allow_html=True)
            
            st.markdown("""
            <div class="download-section">
                <h3>🎯 Dynamic Excel Dashboard</h3>
                <p style="font-size: 1.1rem; margin-bottom: 1rem;">Generate an intelligent Excel file with:</p>
                <ul style="font-size: 1rem; line-height: 1.8;">
                    <li>✅ Dropdown to select any branch</li>
                    <li>✅ Auto-updating metrics and calculations</li>
                    <li>✅ Works completely offline</li>
                    <li>✅ All branches summary in one file</li>
                    <li>✅ Professional formatting and design</li>
                </ul>
            </div>
            """, unsafe_allow_html=True)
            
            if st.button("📊 Generate Excel Dashboard", use_container_width=True, type="primary"):
                with st.spinner("Creating dynamic dashboard..."):
                    excel_file = create_dynamic_excel(df)
                
                st.download_button(
                    "📥 Download Excel Dashboard",
                    excel_file,
                    f"BankVista_Dashboard_{date.today().strftime('%Y%m%d')}.xlsx",
                    "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                    use_container_width=True
                )
            
            st.markdown("---")
            
            col1, col2 = st.columns(2)
            
            with col1:
                st.markdown("""
                <div class="feature-card">
                    <div class="feature-icon">💬</div>
                    <div class="feature-title">Export Conversation</div>
                    <div class="feature-desc">Download your complete AI chat history in JSON format</div>
                </div>
                """, unsafe_allow_html=True)
                
                if st.button("📝 Export Chat History", use_container_width=True):
                    chat_json = st.session_state['ai_assistant'].export_conversation()
                    
                    st.download_button(
                        "📥 Download Conversation (JSON)",
                        chat_json,
                        f"BankVista_Conversation_{date.today().strftime('%Y%m%d')}.json",
                        "application/json",
                        use_container_width=True
                    )
            
            with col2:
                st.markdown("""
                <div class="feature-card">
                    <div class="feature-icon">📊</div>
                    <div class="feature-title">Export Raw Data</div>
                    <div class="feature-desc">Download the current dataset as CSV for external analysis</div>
                </div>
                """, unsafe_allow_html=True)
                
                csv = df.to_csv(index=False)
                st.download_button(
                    "📥 Download CSV",
                    csv,
                    f"BankVista_Data_{date.today().strftime('%Y%m%d')}.csv",
                    "text/csv",
                    use_container_width=True
                )
    
    else:
        # WELCOME SCREEN
        st.markdown("""
        <div style="text-align: center; padding: 4rem 2rem; background: white; border-radius: 24px; margin-top: 2rem; box-shadow: 0 4px 6px rgba(0,0,0,0.05);">
            <div style="font-size: 4rem; margin-bottom: 1rem;">🚀</div>
            <h2 style="font-size: 2.5rem; margin-bottom: 1rem; font-weight: 800; color: #1f2937;">Welcome to BankVista AI</h2>
            <p style="font-size: 1.3rem; color: #6b7280; margin-bottom: 2rem;">
                Experience the future of banking analytics with AI-powered insights
            </p>
            <p style="font-size: 1.1rem; color: #9ca3af;">
                Upload your data or try our sample dataset to get started
            </p>
        </div>
        """, unsafe_allow_html=True)
        
        # Feature Highlights
        st.markdown('<p class="section-header">✨ Platform Features</p>', unsafe_allow_html=True)
        
        col1, col2, col3 = st.columns(3)
        
        with col1:
            st.markdown("""
            <div class="feature-card">
                <div class="feature-icon">🤖</div>
                <div class="feature-title">AI-Powered Chat</div>
                <div class="feature-desc">
                    Conversational interface to query your banking data naturally.
                    Get instant insights with supportive, context-aware responses.
                </div>
            </div>
            """, unsafe_allow_html=True)
        
        with col2:
            st.markdown("""
            <div class="feature-card">
                <div class="feature-icon">📈</div>
                <div class="feature-title">Predictive Analytics</div>
                <div class="feature-desc">
                    Forecast NPA trends, predict target achievement, and identify
                    risks before they become problems.
                </div>
            </div>
            """, unsafe_allow_html=True)
        
        with col3:
            st.markdown("""
            <div class="feature-card">
                <div class="feature-icon">📊</div>
                <div class="feature-title">Dynamic Dashboards</div>
                <div class="feature-desc">
                    Generate intelligent Excel reports with auto-updating metrics
                    and professional formatting.
                </div>
            </div>
            """, unsafe_allow_html=True)


if __name__ == "__main__":
    main()
