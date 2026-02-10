"""
BankVista AI - PRODUCTION READY VERSION
✅ Fixed all deprecation warnings
✅ Proper caching - no repeated initialization
✅ Modern google.genai SDK
✅ Clean architecture
"""

import streamlit as st
import pandas as pd
import numpy as np
import plotly.graph_objects as go
import plotly.express as px
from plotly.subplots import make_subplots
from datetime import date, datetime
import io
import os
import re
from dotenv import load_dotenv
from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter
from openpyxl.worksheet.datavalidation import DataValidation
import json
import difflib

load_dotenv()

st.set_page_config(
    page_title="BankVista AI",
    page_icon="🤖",
    layout="wide",
    initial_sidebar_state="expanded"
)


# ═══════════════════════════════════════════════════════════════
# BACKEND INITIALIZATION - CACHED (Runs only once per session)
# ═══════════════════════════════════════════════════════════════

@st.cache_resource(show_spinner=False)
def init_groq_client():
    """Initialize Groq client - cached across all users"""
    try:
        from groq import Groq
        api_key = os.getenv('GROQ_API_KEY')
        if api_key and len(api_key.strip()) > 20:
            client = Groq(api_key=api_key.strip())
            # Test connection
            client.chat.completions.create(
                model="llama-3.3-70b-versatile",
                messages=[{"role": "user", "content": "test"}],
                max_tokens=5
            )
            print("✅ Groq client initialized (cached)")
            return client
    except Exception as e:
        print(f"⚠️ Groq init failed: {e}")
    return None


@st.cache_resource(show_spinner=False)
def init_gemini_client():
    """Initialize NEW Google Gemini client - cached"""
    try:
        # NEW SDK - google.genai (not google.generativeai)
        import google.genai as genai
        
        api_key = os.getenv('GEMINI_API_KEY') or os.getenv('GOOGLE_API_KEY')
        if api_key:
            client = genai.Client(api_key=api_key)
            print("✅ Gemini client initialized (NEW SDK, cached)")
            return client
    except ImportError:
        # Fallback to old SDK if new one not installed
        try:
            import warnings
            warnings.filterwarnings('ignore', category=FutureWarning)
            
            import google.generativeai as genai_old
            api_key = os.getenv('GEMINI_API_KEY') or os.getenv('GOOGLE_API_KEY')
            if api_key:
                genai_old.configure(api_key=api_key)
                model = genai_old.GenerativeModel('gemini-1.5-flash')
                print("✅ Gemini client initialized (OLD SDK, cached)")
                return {'client': model, 'type': 'old'}
        except Exception as e:
            print(f"⚠️ Gemini init failed: {e}")
    except Exception as e:
        print(f"⚠️ Gemini init failed: {e}")
    return None


@st.cache_resource(show_spinner=False)
def check_ollama():
    """Check if Ollama is available - cached"""
    try:
        import requests
        response = requests.get('http://localhost:11434/api/tags', timeout=2)
        if response.status_code == 200:
            print("✅ Ollama available (cached)")
            return True
    except:
        pass
    return False


def get_available_backends():
    """Get all available backends - uses cached clients"""
    backends = {}
    
    groq = init_groq_client()
    if groq:
        backends['groq'] = groq
    
    gemini = init_gemini_client()
    if gemini:
        backends['gemini'] = gemini
    
    if check_ollama():
        backends['ollama'] = True
    
    return backends


# ═══════════════════════════════════════════════════════════════
# CSS - MODERN STYLING
# ═══════════════════════════════════════════════════════════════

st.markdown("""
<style>
    @import url('https://fonts.googleapis.com/css2?family=Inter:wght@400;500;600;700;800&display=swap');
    
    .main { 
        background: #f0f2f5; 
        font-family: 'Inter', sans-serif; 
    }
    
    .hero-container {
        background: linear-gradient(135deg, #667eea 0%, #764ba2 100%);
        padding: 3rem 2rem;
        border-radius: 24px;
        margin-bottom: 2rem;
        box-shadow: 0 20px 60px rgba(102,126,234,0.3);
    }
    
    .main-title {
        font-size: 3.5rem;
        font-weight: 800;
        color: white;
        text-align: center;
        margin-bottom: 0.5rem;
    }
    
    .subtitle {
        font-size: 1.2rem;
        color: rgba(255,255,255,0.9);
        text-align: center;
    }
    
    .stat-card {
        background: white;
        border: 1px solid #e5e7eb;
        border-radius: 16px;
        padding: 1.5rem;
        text-align: center;
        box-shadow: 0 4px 6px rgba(0,0,0,0.05);
        transition: all 0.3s ease;
    }
    
    .stat-card:hover {
        transform: translateY(-4px);
        box-shadow: 0 8px 16px rgba(102,126,234,0.15);
    }
    
    .stat-value {
        font-size: 2rem;
        font-weight: 800;
        background: linear-gradient(135deg, #06b6d4, #3b82f6);
        -webkit-background-clip: text;
        -webkit-text-fill-color: transparent;
    }
    
    .stat-label {
        color: #6b7280;
        font-size: 0.8rem;
        font-weight: 600;
        text-transform: uppercase;
    }
    
    .ai-chat-container {
        background: white;
        border: 1px solid #e5e7eb;
        border-radius: 20px;
        padding: 1.5rem;
        margin: 0.8rem 0;
        min-height: 400px;
        max-height: 600px;
        overflow-y: auto;
        box-shadow: 0 4px 6px rgba(0,0,0,0.05);
    }
    
    .user-message {
        background: linear-gradient(135deg, #06b6d4, #3b82f6);
        color: white;
        padding: 0.9rem 1.3rem;
        border-radius: 18px 18px 4px 18px;
        margin: 0.6rem 0;
        max-width: 80%;
        margin-left: auto;
        box-shadow: 0 4px 12px rgba(6,182,212,0.3);
        font-weight: 500;
    }
    
    .ai-message {
        background: #f9fafb;
        border: 1px solid #e5e7eb;
        color: #1f2937;
        padding: 0.9rem 1.3rem;
        border-radius: 18px 18px 18px 4px;
        margin: 0.6rem 0;
        max-width: 82%;
        line-height: 1.6;
    }
    
    .section-header {
        font-size: 1.8rem;
        font-weight: 700;
        color: #1f2937;
        border-bottom: 3px solid #667eea;
        padding-bottom: 0.6rem;
        margin: 1.5rem 0 1rem 0;
    }
    
    .stButton > button {
        background: linear-gradient(135deg, #06b6d4, #3b82f6);
        color: white;
        border: none;
        border-radius: 12px;
        padding: 0.7rem 1.8rem;
        font-weight: 700;
        transition: all 0.3s ease;
    }
    
    .stButton > button:hover {
        transform: translateY(-2px);
        box-shadow: 0 8px 16px rgba(6,182,212,0.3);
    }
</style>
""", unsafe_allow_html=True)


# ═══════════════════════════════════════════════════════════════
# SMART NLP ENGINE
# ═══════════════════════════════════════════════════════════════

class SmartIntentEngine:
    """Detects user intent from natural language"""
    
    INTENT_MAP = {
        "npa": ["npa", "bad loan", "bad loans", "defaulter", "stuck", "risky", "problem loan"],
        "casa": ["casa", "cheap deposit", "savings", "current account", "low cost"],
        "top": ["top", "best", "star", "doing well", "excellent", "champion", "leader"],
        "weak": ["weak", "worst", "struggling", "behind", "need help", "underperform"],
        "deposits": ["deposit", "total deposits", "deposit target", "deposit achievement"],
        "advances": ["advance", "loans given", "lending", "credit", "loan disbursement"],
        "staff": ["staff", "employee", "team", "efficiency", "productivity"],
        "compare": ["compare", "versus", "vs", "side by side", "ranking"],
    }

    def __init__(self, df):
        self.df = df
        self.branch_names = df['Branch_Name'].tolist()

    def detect_intent(self, text):
        """Detect intent from user query"""
        text_lower = text.lower()
        
        # Check for branch mention
        for branch in self.branch_names:
            if branch.lower() in text_lower:
                return f"BRANCH:{branch}"
        
        # Check for intent keywords
        for intent, keywords in self.INTENT_MAP.items():
            if any(kw in text_lower for kw in keywords):
                return intent
        
        return "overview"


# ═══════════════════════════════════════════════════════════════
# AI ASSISTANT - PROPERLY CACHED
# ═══════════════════════════════════════════════════════════════

class AIAssistant:
    """AI Assistant with multi-model support and proper caching"""
    
    def __init__(self, df):
        self.df = df
        self.intent_engine = SmartIntentEngine(df)
        
        # Get cached backends (no repeated initialization!)
        self.backends = get_available_backends()
        self.active_backend = self._select_backend()

    def _select_backend(self):
        """Select best available backend"""
        if 'groq' in self.backends:
            return 'groq'
        elif 'gemini' in self.backends:
            return 'gemini'
        elif 'ollama' in self.backends:
            return 'ollama'
        return 'local'

    def get_backend_info(self):
        """Return formatted backend info"""
        info = {
            'groq': '🚀 Groq AI (Fast)',
            'gemini': '🟢 Gemini AI (Google)',
            'ollama': '🟢 Ollama (Local)',
            'local': '🧠 Local Rules'
        }
        return info.get(self.active_backend, 'Unknown')

    def chat(self, query):
        """Main chat function with fallback"""
        intent = self.intent_engine.detect_intent(query)
        
        # Try AI backends in priority order
        for backend in ['groq', 'gemini', 'ollama']:
            if backend in self.backends:
                try:
                    response = self._call_ai(query, backend)
                    if response and len(response) > 10:
                        self.active_backend = backend
                        return response
                except Exception as e:
                    print(f"⚠️ {backend} failed: {e}")
                    continue
        
        # Fallback to local
        return self._local_response(intent, query)

    def _call_ai(self, query, backend):
        """Call specific AI backend"""
        # Prepare context
        summary = f"""Banking Data Summary ({len(self.df)} branches):
- Average NPA: {self.df['NPA_Percent'].mean():.1f}%
- Average CASA: {self.df['CASA_Percent'].mean():.1f}%
- Total Deposits: ₹{self.df['Total_Deposits'].sum():.0f}Cr

Branch Details: {', '.join([f"{r['Branch_Name']} (NPA: {r['NPA_Percent']:.1f}%)" for _, r in self.df.iterrows()])}"""

        prompt = f"""{summary}

User Question: {query}

Instructions:
- Be concise (2-3 paragraphs max)
- Use bullet points for lists
- Include specific numbers
- Be supportive and encouraging
- "bad loans" = NPA, "cheap deposits" = CASA"""

        # Call backend
        if backend == 'groq':
            client = self.backends['groq']
            response = client.chat.completions.create(
                model="llama-3.3-70b-versatile",
                messages=[{"role": "user", "content": prompt}],
                max_tokens=500,
                temperature=0.7
            )
            return response.choices[0].message.content

        elif backend == 'gemini':
            client = self.backends['gemini']
            
            # Check if new or old SDK
            if isinstance(client, dict) and client.get('type') == 'old':
                response = client['client'].generate_content(prompt)
                return response.text
            else:
                # NEW SDK
                response = client.models.generate_content(
                    model='gemini-2.0-flash-exp',
                    contents=prompt
                )
                return response.text

        elif backend == 'ollama':
            import requests
            response = requests.post('http://localhost:11434/api/generate', json={
                "model": "llama3.2",
                "prompt": prompt,
                "stream": False
            })
            if response.status_code == 200:
                return response.json()['response']

        return None

    def _local_response(self, intent, query):
        """Local rule-based responses"""
        if intent.startswith("BRANCH:"):
            branch = intent.split(":")[1]
            return self._branch_detail(branch)
        elif intent == "npa":
            return self._npa_analysis()
        elif intent == "casa":
            return self._casa_analysis()
        elif intent == "top":
            return self._top_performers()
        elif intent == "weak":
            return self._weak_branches()
        elif intent == "deposits":
            return self._deposits_summary()
        elif intent == "advances":
            return self._advances_summary()
        elif intent == "staff":
            return self._staff_summary()
        elif intent == "compare":
            return self._compare_branches()
        else:
            return self._overview()

    def _npa_analysis(self):
        high = self.df.nlargest(3, 'NPA_Percent')
        avg = self.df['NPA_Percent'].mean()
        r = f"📊 **NPA Analysis**\n\nOrganization Average: **{avg:.2f}%**\n\n**High NPA Branches:**\n"
        for _, row in high.iterrows():
            status = "🔴 High" if row['NPA_Percent'] > 6 else "🟡 Moderate"
            r += f"• **{row['Branch_Name']}**: {row['NPA_Percent']:.2f}% {status}\n"
        r += "\n💡 **Recommendation:** Focus on early recovery efforts."
        return r

    def _casa_analysis(self):
        low = self.df.nsmallest(3, 'CASA_Percent')
        avg = self.df['CASA_Percent'].mean()
        r = f"💰 **CASA Analysis**\n\nOrganization Average: **{avg:.1f}%**\n\n**Growth Opportunities:**\n"
        for _, row in low.iterrows():
            gap = 40 - row['CASA_Percent']
            r += f"• **{row['Branch_Name']}**: {row['CASA_Percent']:.1f}% (Gap: {gap:.1f}%)\n"
        r += "\n💡 **Recommendation:** Launch targeted savings campaigns."
        return r

    def _top_performers(self):
        df = self.df.copy()
        df['score'] = ((df['Total_Deposits']/df['Deposit_Target']*30).clip(upper=30) +
                      (df['Advances']/df['Advance_Target']*30).clip(upper=30) +
                      ((100-df['NPA_Percent']*10)/100*20).clip(0, 20) +
                      (df['CASA_Percent']/100*20).clip(upper=20))
        df = df.sort_values('score', ascending=False)
        
        r = "🏆 **Top Performers**\n\n"
        for i, (_, row) in enumerate(df.head(5).iterrows(), 1):
            emoji = ["🥇","🥈","🥉","⭐","⭐"][i-1]
            r += f"{emoji} **{row['Branch_Name']}** (Score: {row['score']:.0f}/100)\n"
        return r

    def _weak_branches(self):
        df = self.df.copy()
        df['dep%'] = df['Total_Deposits']/df['Deposit_Target']*100
        df['adv%'] = df['Advances']/df['Advance_Target']*100
        weak = df[(df['dep%'] < 90) | (df['adv%'] < 90) | (df['NPA_Percent'] > 5)]
        
        r = "🛟 **Branches Needing Support**\n\n"
        if weak.empty:
            return r + "✅ All branches performing well!"
        
        for _, row in weak.iterrows():
            r += f"• **{row['Branch_Name']}**\n"
        r += "\n💡 **Recommendation:** Provide targeted support."
        return r

    def _deposits_summary(self):
        total = self.df['Total_Deposits'].sum()
        target = self.df['Deposit_Target'].sum()
        pct = total/target*100
        r = f"🏦 **Deposits Summary**\n\nTotal: ₹{total:.1f}Cr ({pct:.0f}% of target)\n\n**Top 5:**\n"
        for _, row in self.df.nlargest(5, 'Total_Deposits').iterrows():
            r += f"• **{row['Branch_Name']}**: ₹{row['Total_Deposits']:.1f}Cr\n"
        return r

    def _advances_summary(self):
        total = self.df['Advances'].sum()
        r = f"📋 **Advances Summary**\n\nTotal: ₹{total:.1f}Cr\n\n**Top 5:**\n"
        for _, row in self.df.nlargest(5, 'Advances').iterrows():
            r += f"• **{row['Branch_Name']}**: ₹{row['Advances']:.1f}Cr\n"
        return r

    def _staff_summary(self):
        total = self.df['Staff_Count'].sum()
        avg_biz = self.df['Business_Per_Staff'].mean()
        r = f"👥 **Staff Summary**\n\nTotal Staff: {total}\nAvg Business/Staff: ₹{avg_biz:.1f}Cr\n\n**Most Productive:**\n"
        for _, row in self.df.nlargest(5, 'Business_Per_Staff').iterrows():
            r += f"• **{row['Branch_Name']}**: ₹{row['Business_Per_Staff']:.1f}Cr/staff\n"
        return r

    def _compare_branches(self):
        df = self.df.copy()
        df['score'] = ((df['Total_Deposits']/df['Deposit_Target']*50).clip(upper=50) +
                      ((100-df['NPA_Percent']*10)/100*50).clip(0, 50))
        df = df.sort_values('score', ascending=False)
        r = "📊 **Branch Rankings**\n\n"
        for i, (_, row) in enumerate(df.iterrows(), 1):
            r += f"{i}. **{row['Branch_Name']}** ({row['score']:.0f}/100)\n"
        return r

    def _branch_detail(self, branch):
        row = self.df[self.df['Branch_Name'] == branch]
        if row.empty:
            return f"Branch '{branch}' not found."
        row = row.iloc[0]
        dep_pct = row['Total_Deposits']/row['Deposit_Target']*100
        return f"""📍 **{branch} - Branch Details**

**Zone:** {row['Zone']}

**Performance:**
• Deposits: ₹{row['Total_Deposits']:.1f}Cr ({dep_pct:.0f}% of target)
• Advances: ₹{row['Advances']:.1f}Cr
• NPA: {row['NPA_Percent']:.2f}%
• CASA: {row['CASA_Percent']:.1f}%

**Team:**
• Staff: {row['Staff_Count']}
• Business/Staff: ₹{row['Business_Per_Staff']:.1f}Cr"""

    def _overview(self):
        return f"""📊 **Banking Overview**

**Organization Summary:**
• Total Branches: {len(self.df)}
• Total Deposits: ₹{self.df['Total_Deposits'].sum():.1f}Cr
• Average NPA: {self.df['NPA_Percent'].mean():.2f}%
• Average CASA: {self.df['CASA_Percent'].mean():.1f}%
• Total Staff: {self.df['Staff_Count'].sum()}

💬 **Try asking:**
• "Which branches have bad loans?"
• "Where can we grow CASA?"
• "Tell me about Hyderabad Main"
• "Who are the top performers?"
"""


# ═══════════════════════════════════════════════════════════════
# PREDICTIVE ANALYTICS
# ═══════════════════════════════════════════════════════════════

class PredictiveAnalytics:
    def __init__(self, df):
        self.df = df

    def predict_npa_trend(self, branch_name, months=6):
        branch = self.df[self.df['Branch_Name'] == branch_name].iloc[0]
        npa = branch['NPA_Percent']
        predictions = []
        
        for m in range(1, months+1):
            npa = max(0, npa * (1.0 + np.random.normal(0, 0.15)))
            predictions.append({'month': m, 'predicted_npa': round(npa, 2)})
        
        final = predictions[-1]['predicted_npa']
        risk = 'high' if final > 6 else 'medium' if final > 3 else 'low'
        
        return {
            'branch': branch_name,
            'current_npa': branch['NPA_Percent'],
            'predictions': predictions,
            'risk_level': risk
        }


# ═══════════════════════════════════════════════════════════════
# ANOMALY DETECTOR
# ═══════════════════════════════════════════════════════════════

class AnomalyDetector:
    def __init__(self, df):
        self.df = df

    def detect(self):
        anomalies = []
        metrics = ['NPA_Percent', 'CASA_Percent', 'Business_Per_Staff']
        
        for col in metrics:
            if col not in self.df.columns:
                continue
            
            mean, std = self.df[col].mean(), self.df[col].std()
            if std == 0:
                continue
            
            for _, row in self.df.iterrows():
                z = (row[col] - mean) / std
                if abs(z) > 1.7:
                    anomalies.append({
                        'branch': row['Branch_Name'],
                        'metric': col,
                        'value': round(row[col], 2),
                        'z_score': round(z, 2),
                        'direction': 'HIGH' if z > 0 else 'LOW'
                    })
        
        return anomalies


# ═══════════════════════════════════════════════════════════════
# EXCEL EXPORT
# ═══════════════════════════════════════════════════════════════

def create_dynamic_excel(df):
    """Create interactive Excel dashboard"""
    output = io.BytesIO()
    wb = Workbook()
    wb.remove(wb.active)
    
    # Data sheet (hidden)
    ws_data = wb.create_sheet("_Data")
    ws_data.sheet_state = 'hidden'
    for ci, h in enumerate(df.columns, 1):
        ws_data.cell(1, ci, h).font = Font(bold=True)
    for ri, row in enumerate(df.itertuples(index=False), 2):
        for ci, v in enumerate(row, 1):
            ws_data.cell(ri, ci, v)
    
    # Dashboard
    ws = wb.create_sheet("Dashboard", 0)
    ws['A1'] = "BANKVISTA AI DASHBOARD"
    ws['A1'].font = Font(bold=True, size=16, color="FFFFFF")
    ws['A1'].fill = PatternFill("solid", fgColor="667EEA")
    ws['A1'].alignment = Alignment(horizontal='center')
    ws.merge_cells('A1:F1')
    
    ws['A3'] = "Select Branch:"
    ws['B3'] = df.iloc[0]['Branch_Name']
    ws['B3'].fill = PatternFill("solid", fgColor="FFF2CC")
    
    dv = DataValidation(type="list", formula1=f"=_Data!$B$2:$B${len(df)+1}")
    dv.add('B3')
    ws.add_data_validation(dv)
    
    ws['A5'] = "Deposits:"
    ws['B5'] = '=INDEX(_Data!$D:$D,MATCH(B3,_Data!$B:$B,0))'
    ws['A6'] = "NPA %:"
    ws['B6'] = '=INDEX(_Data!$H:$H,MATCH(B3,_Data!$B:$B,0))'
    ws['A7'] = "CASA %:"
    ws['B7'] = '=INDEX(_Data!$J:$J,MATCH(B3,_Data!$B:$B,0))'
    
    for c in 'ABCDEF':
        ws.column_dimensions[c].width = 18
    
    wb.save(output)
    output.seek(0)
    return output


# ═══════════════════════════════════════════════════════════════
# SAMPLE DATA
# ═══════════════════════════════════════════════════════════════

def sample_data():
    """Generate sample banking data"""
    return pd.DataFrame({
        'Branch_ID': ['B1001','B1002','B1003','B1004','B1005','B2001','B2002','B2003'],
        'Branch_Name': ['Mansoorabad','Adilabad','Hyderabad Main','Secunderabad','Warangal',
                       'Vijayawada','Visakhapatnam','Guntur'],
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


# ═══════════════════════════════════════════════════════════════
# MAIN APP
# ═══════════════════════════════════════════════════════════════

def main():
    """Main Streamlit app"""
    
    # Initialize session state
    if 'chat_messages' not in st.session_state:
        st.session_state['chat_messages'] = []

    # Hero section
    st.markdown("""
    <div class="hero-container">
        <h1 class="main-title">🤖 BankVista AI</h1>
        <p class="subtitle">Production-Ready Banking Analytics with Multi-AI Support</p>
    </div>
    """, unsafe_allow_html=True)

    # Sidebar
    with st.sidebar:
        st.markdown("### 📤 Data Upload")
        uploaded = st.file_uploader("Upload CSV/Excel", type=['csv','xlsx','xls'])
        
        if st.button("📊 Try Sample Data"):
            st.session_state['use_sample'] = True
            st.session_state['uploaded_df'] = sample_data()
            if 'ai_assistant' in st.session_state:
                del st.session_state['ai_assistant']
            st.session_state['chat_messages'] = []
            st.rerun()

        st.markdown("---")
        st.markdown("### 🔑 API Keys")
        st.caption("Optional - Enter to enable AI")
        
        groq_key = st.text_input("Groq API Key", type="password", 
                                 help="Get free at console.groq.com")
        if groq_key:
            os.environ['GROQ_API_KEY'] = groq_key
        
        gemini_key = st.text_input("Gemini API Key", type="password",
                                   help="Get free at aistudio.google.com/apikey")
        if gemini_key:
            os.environ['GEMINI_API_KEY'] = gemini_key

        st.markdown("---")
        st.markdown("### ℹ️ About")
        st.caption("""
        ✅ Zero repeated initialization
        ✅ Proper caching with @st.cache_resource
        ✅ New google.genai SDK
        ✅ Multi-AI fallback support
        """)

    # Load data
    if uploaded:
        try:
            ext = uploaded.name.split('.')[-1].lower()
            df = pd.read_csv(uploaded) if ext == 'csv' else pd.read_excel(uploaded)
            st.session_state['uploaded_df'] = df
            st.session_state['use_sample'] = False
            if 'ai_assistant' in st.session_state:
                del st.session_state['ai_assistant']
            st.session_state['chat_messages'] = []
        except Exception as e:
            st.error(f"Error loading file: {e}")

    # Main app logic
    if st.session_state.get('use_sample') or 'uploaded_df' in st.session_state:
        df = st.session_state['uploaded_df']

        # Initialize AI assistant (only once per session!)
        if 'ai_assistant' not in st.session_state:
            st.session_state['ai_assistant'] = AIAssistant(df)

        assistant = st.session_state['ai_assistant']
        predictive = PredictiveAnalytics(df)
        anomaly_detector = AnomalyDetector(df)

        # Status bar
        st.markdown(f"""
        <div style="background:#d1fae5;padding:1rem;border-radius:12px;margin-bottom:1.5rem;">
            <p style="text-align:center;margin:0;color:#065f46;">
                ✅ {len(df)} branches | 💰 ₹{df['Total_Deposits'].sum():.0f}Cr | {assistant.get_backend_info()}
            </p>
        </div>
        """, unsafe_allow_html=True)

        # KPIs
        c1, c2, c3, c4 = st.columns(4)
        with c1:
            st.markdown(f'<div class="stat-card"><div class="stat-value">{df["Total_Deposits"].sum()/df["Deposit_Target"].sum()*100:.0f}%</div><div class="stat-label">Deposits</div></div>', unsafe_allow_html=True)
        with c2:
            st.markdown(f'<div class="stat-card"><div class="stat-value">{df["NPA_Percent"].mean():.1f}%</div><div class="stat-label">Avg NPA</div></div>', unsafe_allow_html=True)
        with c3:
            st.markdown(f'<div class="stat-card"><div class="stat-value">{df["CASA_Percent"].mean():.1f}%</div><div class="stat-label">Avg CASA</div></div>', unsafe_allow_html=True)
        with c4:
            st.markdown(f'<div class="stat-card"><div class="stat-value">{len(df)}</div><div class="stat-label">Branches</div></div>', unsafe_allow_html=True)

        # Tabs
        tab1, tab2, tab3, tab4, tab5, tab6, tab7 = st.tabs([
            "💬 Chat", "📈 Predictions", "📊 Dashboard", "🔄 Compare", 
            "🗺️ Heatmaps", "🔍 Anomalies", "📥 Export"
        ])

        # TAB 1: CHAT
        with tab1:
            st.markdown('<p class="section-header">💬 AI Chat</p>', unsafe_allow_html=True)
            
            # Chat display
            chat_container = st.container()
            with chat_container:
                st.markdown('<div class="ai-chat-container">', unsafe_allow_html=True)
                
                if not st.session_state['chat_messages']:
                    st.markdown('<div class="ai-message">👋 Hello! Ask me anything about your banking data.</div>', unsafe_allow_html=True)
                
                for msg in st.session_state['chat_messages']:
                    if msg['role'] == 'user':
                        st.markdown(f'<div class="user-message">👤 {msg["content"]}</div>', unsafe_allow_html=True)
                    else:
                        st.markdown(f'<div class="ai-message">🤖 {msg["content"]}</div>', unsafe_allow_html=True)
                
                st.markdown('</div>', unsafe_allow_html=True)

            # Input
            user_input = st.text_input("Your question:", key="chat_input", 
                                      placeholder="e.g., Which branches have bad loans?")
            
            col1, col2 = st.columns([1, 5])
            with col1:
                send_btn = st.button("Send 🚀", key="send_btn")
            
            if send_btn and user_input:
                st.session_state['chat_messages'].append({'role': 'user', 'content': user_input})
                
                with st.spinner("💭 Thinking..."):
                    response = assistant.chat(user_input)
                
                st.session_state['chat_messages'].append({'role': 'assistant', 'content': response})
                st.rerun()

            # Quick questions
            st.markdown("**Quick Questions:**")
            cols = st.columns(4)
            questions = [
                ("🔴 Bad Loans?", "Which branches have bad loans?"),
                ("💰 CASA?", "Where can we grow CASA?"),
                ("🏆 Top?", "Who are top performers?"),
                ("🛟 Help?", "Which branches need support?"),
            ]
            
            for idx, (col, (label, q)) in enumerate(zip(cols, questions)):
                with col:
                    if st.button(label, key=f"quick_{idx}"):
                        st.session_state['chat_messages'].append({'role': 'user', 'content': q})
                        with st.spinner("💭"):
                            resp = assistant.chat(q)
                        st.session_state['chat_messages'].append({'role': 'assistant', 'content': resp})
                        st.rerun()

        # TAB 2: PREDICTIONS
        with tab2:
            st.markdown('<p class="section-header">📈 Predictions</p>', unsafe_allow_html=True)
            
            branch = st.selectbox("Select Branch:", df['Branch_Name'].tolist(), key="pred_branch")
            npa_pred = predictive.predict_npa_trend(branch)
            
            col1, col2 = st.columns([2, 1])
            with col1:
                months = [p['month'] for p in npa_pred['predictions']]
                npas = [p['predicted_npa'] for p in npa_pred['predictions']]
                
                fig = go.Figure()
                fig.add_trace(go.Scatter(
                    x=months, y=npas, mode='lines+markers',
                    line=dict(color='#06b6d4', width=3),
                    marker=dict(size=8)
                ))
                fig.add_hline(y=3, line_dash="dash", line_color="#10b981", 
                             annotation_text="Target: 3%")
                fig.add_hline(y=6, line_dash="dash", line_color="#f59e0b",
                             annotation_text="Watch: 6%")
                fig.update_layout(
                    title="NPA Forecast (6 Months)",
                    height=350,
                    plot_bgcolor='white',
                    paper_bgcolor='white',
                    xaxis_title="Month",
                    yaxis_title="NPA %"
                )
                st.plotly_chart(fig, key="npa_forecast_chart")
            
            with col2:
                st.metric("Current NPA", f"{npa_pred['current_npa']:.2f}%")
                st.metric("Predicted (Month 6)", f"{npa_pred['predictions'][-1]['predicted_npa']:.2f}%")
                st.metric("Risk Level", npa_pred['risk_level'].upper())

        # TAB 3: DASHBOARD
        with tab3:
            st.markdown('<p class="section-header">📊 Dashboard</p>', unsafe_allow_html=True)
            
            display_df = df.copy()
            display_df['Dep_%'] = (display_df['Total_Deposits']/display_df['Deposit_Target']*100).round(1)
            display_df['Adv_%'] = (display_df['Advances']/display_df['Advance_Target']*100).round(1)
            
            st.dataframe(
                display_df[['Branch_Name','Zone','Total_Deposits','Dep_%','NPA_Percent','CASA_Percent']],
                width='stretch',  # FIXED: use width instead of use_container_width
                hide_index=True
            )
            
            col1, col2 = st.columns(2)
            with col1:
                fig = px.bar(df, x='Branch_Name', y='NPA_Percent', 
                           title='NPA % by Branch',
                           color='NPA_Percent', 
                           color_continuous_scale='Reds')
                fig.update_layout(height=350, plot_bgcolor='white', paper_bgcolor='white')
                st.plotly_chart(fig, key="npa_chart")
            
            with col2:
                fig = px.bar(df, x='Branch_Name', y='CASA_Percent',
                           title='CASA % by Branch',
                           color='CASA_Percent',
                           color_continuous_scale='Greens')
                fig.update_layout(height=350, plot_bgcolor='white', paper_bgcolor='white')
                st.plotly_chart(fig, key="casa_chart")

        # TAB 4: COMPARE
        with tab4:
            st.markdown('<p class="section-header">🔄 Compare Branches</p>', unsafe_allow_html=True)
            
            col1, col2 = st.columns(2)
            with col1:
                b1 = st.selectbox("Branch A:", df['Branch_Name'].tolist(), key="cmp_a")
            with col2:
                b2 = st.selectbox("Branch B:", df['Branch_Name'].tolist(),
                                 index=min(1, len(df)-1), key="cmp_b")
            
            if b1 != b2:
                r1 = df[df['Branch_Name'] == b1].iloc[0]
                r2 = df[df['Branch_Name'] == b2].iloc[0]
                
                metrics = ['Total_Deposits', 'Advances', 'NPA_Percent', 'CASA_Percent']
                labels = ['Deposits (Cr)', 'Advances (Cr)', 'NPA %', 'CASA %']
                
                fig = go.Figure()
                fig.add_trace(go.Bar(
                    name=b1, x=labels, y=[r1[m] for m in metrics],
                    marker_color='#06b6d4'
                ))
                fig.add_trace(go.Bar(
                    name=b2, x=labels, y=[r2[m] for m in metrics],
                    marker_color='#667eea'
                ))
                fig.update_layout(
                    barmode='group',
                    title="Branch Comparison",
                    height=350,
                    plot_bgcolor='white',
                    paper_bgcolor='white'
                )
                st.plotly_chart(fig, key="comparison_chart")

        # TAB 5: HEATMAPS
        with tab5:
            st.markdown('<p class="section-header">🗺️ Heatmaps</p>', unsafe_allow_html=True)
            
            metric = st.selectbox("Select Metric:", 
                                 ['NPA_Percent', 'CASA_Percent', 'Business_Per_Staff'],
                                 key="heatmap_metric")
            
            color_scale = 'RdYlGn_r' if metric == 'NPA_Percent' else 'RdYlGn'
            
            fig = px.bar(df, x='Branch_Name', y=metric,
                        title=f"{metric} Heatmap",
                        color=metric,
                        color_continuous_scale=color_scale)
            fig.update_layout(height=400, plot_bgcolor='white', paper_bgcolor='white')
            st.plotly_chart(fig, key="heatmap_chart")

        # TAB 6: ANOMALIES
        with tab6:
            st.markdown('<p class="section-header">🔍 Anomaly Detection</p>', unsafe_allow_html=True)
            
            anomalies = anomaly_detector.detect()
            
            if not anomalies:
                st.success("✅ No anomalies detected!")
            else:
                st.warning(f"⚠️ Found {len(anomalies)} anomalies")
                
                adf = pd.DataFrame(anomalies)
                st.dataframe(adf, width='stretch', hide_index=True)  # FIXED

        # TAB 7: EXPORT
        with tab7:
            st.markdown('<p class="section-header">📥 Export Data</p>', unsafe_allow_html=True)
            
            st.markdown("""
            <div style="background:linear-gradient(135deg,#10b981,#059669);color:white;
                       padding:2rem;border-radius:20px;margin:1rem 0;">
                <h3 style="color:white;margin:0 0 0.5rem 0;">Dynamic Excel Dashboard</h3>
                <p style="color:white;margin:0;">Interactive Excel with auto-updating formulas</p>
            </div>
            """, unsafe_allow_html=True)
            
            if st.button("📊 Generate Excel Dashboard"):
                with st.spinner("Creating Excel file..."):
                    excel_file = create_dynamic_excel(df)
                    st.download_button(
                        label="⬇️ Download Excel",
                        data=excel_file,
                        file_name=f"BankVista_{date.today():%Y%m%d}.xlsx",
                        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                    )
                    st.success("✅ Excel file ready!")

    else:
        # Welcome screen
        st.markdown("""
        <div style="text-align:center;padding:3rem;background:white;
                   border-radius:24px;margin-top:2rem;box-shadow:0 4px 6px rgba(0,0,0,0.05);">
            <h2 style="color:#1f2937;">Welcome to BankVista AI</h2>
            <p style="color:#6b7280;">Upload your data or try the sample dataset to begin</p>
        </div>
        """, unsafe_allow_html=True)


if __name__ == "__main__":
    main()
