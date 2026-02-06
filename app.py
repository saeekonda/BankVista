"""
BankVista AI - Enhanced Edition
- Smart NLP Intent Layer (understands layman terms)
- Follow-up context & entity memory
- Branch Comparison, Heatmap, Anomaly Detection
- Staff Efficiency, Zone Analytics, Target Gap Tracker
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
    page_title="BankVista AI - Intelligent Banking Analytics",
    page_icon="🤖",
    layout="wide",
    initial_sidebar_state="expanded"
)

# ─────────────────────────────────────────────
# CSS
# ─────────────────────────────────────────────
st.markdown("""
<style>
    @import url('https://fonts.googleapis.com/css2?family=Inter:wght@400;500;600;700;800&family=Space+Grotesk:wght@500;600;700;800&display=swap');

    .main { background: #f0f2f5; font-family: 'Inter', sans-serif; }

    /* Hero */
    .hero-container {
        background: linear-gradient(135deg, #667eea 0%, #764ba2 100%);
        padding: 3.5rem 2rem; border-radius: 24px; margin-bottom: 2rem;
        box-shadow: 0 20px 60px rgba(102,126,234,0.3);
        position: relative; overflow: hidden;
    }
    .hero-container::before {
        content:''; position:absolute; top:-50%; right:-50%;
        width:200%; height:200%;
        background: radial-gradient(circle, rgba(255,255,255,0.08) 1px, transparent 1px);
        background-size: 50px 50px;
        animation: grid-move 20s linear infinite;
    }
    @keyframes grid-move { 0%{transform:translate(0,0)} 100%{transform:translate(50px,50px)} }

    .main-title {
        font-family:'Space Grotesk',sans-serif; font-size:3.8rem; font-weight:800;
        color:white; text-align:center; margin-bottom:0.8rem;
        position:relative; z-index:1; text-shadow:0 2px 20px rgba(0,0,0,0.2); letter-spacing:-2px;
    }
    .subtitle { font-size:1.3rem; color:rgba(255,255,255,0.92); text-align:center; font-weight:500; margin-bottom:1rem; position:relative; z-index:1; }
    .feature-pills { display:flex; flex-wrap:wrap; justify-content:center; gap:0.8rem; margin-top:1.5rem; position:relative; z-index:1; }
    .feature-pill {
        background:rgba(255,255,255,0.18); backdrop-filter:blur(10px);
        border:1px solid rgba(255,255,255,0.3); padding:0.5rem 1.1rem;
        border-radius:25px; color:white; font-weight:600; font-size:0.85rem;
        display:inline-flex; align-items:center; gap:0.4rem; transition:all 0.3s ease;
    }
    .feature-pill:hover { background:rgba(255,255,255,0.28); transform:translateY(-2px); }

    /* Cards */
    .stat-card {
        background:white; border:1px solid #e5e7eb; border-radius:16px;
        padding:1.6rem; text-align:center; transition:all 0.3s ease;
        box-shadow:0 4px 6px rgba(0,0,0,0.05);
    }
    .stat-card:hover { transform:translateY(-4px); box-shadow:0 10px 30px rgba(102,126,234,0.15); border-color:#667eea; }
    .stat-value {
        font-size:2.2rem; font-weight:800;
        background:linear-gradient(135deg,#06b6d4,#3b82f6);
        -webkit-background-clip:text; -webkit-text-fill-color:transparent;
        margin-bottom:0.4rem;
    }
    .stat-label { color:#6b7280; font-size:0.82rem; font-weight:600; text-transform:uppercase; letter-spacing:1px; }
    .stat-trend { color:#10b981; font-size:0.8rem; font-weight:600; margin-top:0.4rem; }

    .section-header {
        font-family:'Space Grotesk',sans-serif; font-size:1.8rem; font-weight:700;
        color:#1f2937; border-bottom:3px solid #667eea; padding-bottom:0.6rem; margin:1.5rem 0 1rem 0;
    }

    .feature-card {
        background:white; border:1px solid #e5e7eb; border-radius:20px;
        padding:1.8rem; margin:0.8rem 0; transition:all 0.3s ease;
        box-shadow:0 4px 6px rgba(0,0,0,0.05);
    }
    .feature-card:hover { transform:translateY(-4px); box-shadow:0 12px 35px rgba(102,126,234,0.18); border-color:#667eea; }
    .feature-icon { font-size:2.5rem; margin-bottom:0.8rem; }
    .feature-title { font-size:1.3rem; font-weight:700; color:#1f2937; margin-bottom:0.4rem; }
    .feature-desc { color:#6b7280; font-size:0.95rem; line-height:1.6; }

    /* Chat */
    .ai-chat-container {
        background:white; border:1px solid #e5e7eb; border-radius:20px;
        padding:1.5rem; margin:0.8rem 0; min-height:300px; max-height:520px;
        overflow-y:auto; box-shadow:0 4px 6px rgba(0,0,0,0.05);
    }
    .user-message {
        background:linear-gradient(135deg,#06b6d4,#3b82f6); color:white;
        padding:0.9rem 1.3rem; border-radius:18px 18px 4px 18px;
        margin:0.6rem 0; max-width:80%; margin-left:auto;
        box-shadow:0 4px 12px rgba(6,182,212,0.3); font-weight:500;
    }
    .ai-message {
        background:#f9fafb; border:1px solid #e5e7eb; color:#1f2937;
        padding:0.9rem 1.3rem; border-radius:18px 18px 18px 4px;
        margin:0.6rem 0; max-width:82%; box-shadow:0 2px 4px rgba(0,0,0,0.05); line-height:1.6;
    }
    .typing-indicator { color:#9ca3af; font-style:italic; padding:0.5rem 1rem; }

    /* Alerts */
    .alert-info { background:#dbeafe; border-left:4px solid #3b82f6; padding:1.2rem; margin:0.6rem 0; border-radius:8px; color:#1e40af; }
    .alert-watch { background:#fef3c7; border-left:4px solid #f59e0b; padding:1.2rem; margin:0.6rem 0; border-radius:8px; color:#92400e; }
    .alert-strength { background:#d1fae5; border-left:4px solid #10b981; padding:1.2rem; margin:0.6rem 0; border-radius:8px; color:#065f46; }
    .alert-anomaly { background:#fce7f3; border-left:4px solid #ec4899; padding:1.2rem; margin:0.6rem 0; border-radius:8px; color:#831843; }

    /* Buttons */
    .stButton > button {
        background:linear-gradient(135deg,#06b6d4,#3b82f6); color:white; border:none;
        border-radius:12px; padding:0.7rem 1.8rem; font-weight:700; font-size:0.95rem;
        transition:all 0.3s ease; box-shadow:0 4px 12px rgba(6,182,212,0.3);
    }
    .stButton > button:hover { transform:translateY(-2px); box-shadow:0 6px 18px rgba(6,182,212,0.4); }

    .download-section {
        background:linear-gradient(135deg,#10b981,#059669); color:white;
        padding:2rem; border-radius:20px; margin:1.5rem 0;
        box-shadow:0 10px 30px rgba(16,185,129,0.3);
    }
    .download-section h3 { font-size:1.6rem; font-weight:700; margin-bottom:0.8rem; color:white; }
    .download-section p, .download-section li { color:white; }

    .metric-card {
        background:white; border:1px solid #e5e7eb; padding:1.3rem;
        border-radius:16px; box-shadow:0 4px 6px rgba(0,0,0,0.05);
        margin:0.4rem 0; transition:all 0.3s ease;
    }
    .metric-card:hover { transform:translateY(-2px); box-shadow:0 8px 25px rgba(102,126,234,0.15); }
    .metric-card p { color:#1f2937; }

    /* Compare */
    .compare-card {
        background:white; border:2px solid #e5e7eb; border-radius:16px;
        padding:1.5rem; transition:all 0.3s ease;
    }
    .compare-card:hover { border-color:#667eea; box-shadow:0 8px 25px rgba(102,126,234,0.15); }
    .compare-header { font-weight:700; font-size:1.2rem; color:#1f2937; border-bottom:2px solid #667eea; padding-bottom:0.5rem; margin-bottom:1rem; }
    .compare-row { display:flex; justify-content:space-between; padding:0.5rem 0; border-bottom:1px solid #f3f4f6; }
    .compare-row:last-child { border-bottom:none; }
    .compare-label { color:#6b7280; font-weight:500; }
    .compare-value { font-weight:700; color:#1f2937; }

    /* Heatmap legend */
    .heatmap-legend { display:flex; gap:1rem; justify-content:center; margin:0.8rem 0; flex-wrap:wrap; }
    .legend-item { display:flex; align-items:center; gap:0.3rem; font-size:0.82rem; font-weight:600; color:#374151; }
    .legend-swatch { width:22px; height:22px; border-radius:4px; }

    /* Tabs */
    .stTabs [data-baseweb="tab-list"] { gap:6px; background:white; padding:0.4rem; border-radius:12px; border:1px solid #e5e7eb; }
    .stTabs [data-baseweb="tab"] { background:transparent; border-radius:8px; color:#6b7280; font-weight:600; padding:0.65rem 1.2rem; }
    .stTabs [aria-selected="true"] { background:linear-gradient(135deg,#06b6d4,#3b82f6); color:white !important; }

    /* Inputs */
    .stTextInput > div > div > input { background:white; border:1px solid #d1d5db; border-radius:12px; color:#1f2937; padding:0.7rem; }
    .stTextInput > div > div > input:focus { border-color:#06b6d4; box-shadow:0 0 0 3px rgba(6,182,212,0.1); }
    .stSelectbox > div > div { background:white; border:1px solid #d1d5db; border-radius:12px; }

    /* Sidebar */
    section[data-testid="stSidebar"] { background:#f9fafb; }

    /* Global text */
    h1,h2,h3,h4,h5,h6 { color:#1f2937 !important; }
    p, span, div, label { color:#374151 !important; }
    .stMarkdown { color:#374151 !important; }

    ::-webkit-scrollbar { width:8px; height:8px; }
    ::-webkit-scrollbar-track { background:#f1f1f1; border-radius:4px; }
    ::-webkit-scrollbar-thumb { background:linear-gradient(135deg,#06b6d4,#3b82f6); border-radius:4px; }
</style>
""", unsafe_allow_html=True)


# ─────────────────────────────────────────────
# SMART NLP / INTENT ENGINE
# ─────────────────────────────────────────────
class SmartIntentEngine:
    """
    Maps layman / casual terms → banking intents.
    Handles typos, partial branch names, follow-ups, and implicit references.
    """

    # Synonym dictionary: intent → list of casual/layman phrases
    INTENT_MAP = {
        "npa": [
            "npa", "bad loan", "bad loans", "defaulter", "defaulters", "default",
            "non performing", "non-performing", "risky loan", "risky loans",
            "problematic loan", "stuck loan", "stuck loans", "recovery",
            "which loans are bad", "loan problems", "loan trouble", "troubled loans",
            "which branches have problems", "problem branches", "red flag",
            "red flags", "worry", "worrying", "concerning", "high risk",
            "danger", "dangerous loans", "bad debt", "loss", "losses",
            "unpaid loans", "overdue", "delayed payment", "delayed payments",
            "loan default", "what loans are stuck", "bad assets"
        ],
        "casa": [
            "casa", "current account", "savings account", "current & savings",
            "cheap deposits", "low cost deposits", "low-cost deposits",
            "current savings", "where can we get more deposits",
            "how to increase deposits", "deposit growth", "deposit opportunities",
            "savings", "current account savings", "casa ratio",
            "low cost fund", "sticky deposits", "deposit mix"
        ],
        "compare": [
            "compare", "comparison", "versus", "vs", "side by side",
            "how do branches compare", "which is better", "better branch",
            "best branch", "top branch", "top branches", "ranking",
            "rank", "who is leading", "leader", "leaders",
            "best performing", "top performing", "compare branches",
            "pit against", "head to head"
        ],
        "top_performers": [
            "top", "best", "star", "star performers", "success", "successful",
            "who is doing well", "doing great", "excellent branches",
            "champions", "leaders", "winner", "winners",
            "high performers", "what are they doing right", "learn from best"
        ],
        "weak_branches": [
            "weak", "weakest", "worst", "struggling", "behind",
            "lagging", "need help", "need support", "underperforming",
            "low performers", "who needs help", "which branch is behind",
            "falling behind", "not meeting targets", "below target"
        ],
        "deposits": [
            "deposit", "deposits", "total deposits", "how much deposits",
            "deposit target", "deposit achievement", "deposit performance",
            "money in bank", "how much money", "collection"
        ],
        "advances": [
            "advance", "advances", "loans given", "loan disbursement",
            "credit", "lending", "how much loans", "advance target",
            "loan performance", "loan achievement"
        ],
        "staff": [
            "staff", "employees", "team", "how many people", "staff efficiency",
            "productivity", "profit per staff", "business per staff",
            "per employee", "team performance", "workforce"
        ],
        "target": [
            "target", "goal", "goals", "achievement", "how close to target",
            "are we meeting target", "target gap", "gap", "how far from target",
            "target progress", "progress", "on track", "are we on track"
        ],
        "anomaly": [
            "anomaly", "anomalies", "odd", "strange", "unusual",
            "something wrong", "anything weird", "outlier", "outliers",
            "abnormal", "unexpected", "flag", "flagged", "suspicious"
        ],
        "zone": [
            "zone", "zones", "region", "regions", "telangana", "andhra",
            "ap", "zone wise", "zone-wise", "by zone", "region wise"
        ],
        "overview": [
            "overview", "summary", "give me summary", "what's going on",
            "status", "how things are", "overall", "big picture",
            "at a glance", "dashboard", "general", "how are we doing",
            "health", "health check", "how is bank doing"
        ],
        "help": [
            "help", "what can you do", "features", "menu", "options",
            "commands", "how to use", "guide", "instructions"
        ],
    }

    FOLLOW_UP_PATTERNS = [
        r"^(tell me more|elaborate|explain more|more details|go on|continue|say more|detail)$",
        r"^(yes|yeah|yep|ok|okay|sure|go ahead|please)$",
        r"^(what about (that|this|it|them|him|her))$",
        r"^(and (that|this|the other one))$",
        r"^(which (one|branch))$",
        r"^(show me|show)$",
    ]

    def __init__(self, df):
        self.df = df
        self.branch_names = df['Branch_Name'].tolist()
        self.branch_names_lower = [n.lower() for n in self.branch_names]

    # ── Public API ──
    def detect_intent(self, text: str) -> str:
        text_clean = text.strip().lower()

        # 1. Check follow-up
        if self._is_follow_up(text_clean):
            return "FOLLOW_UP"

        # 2. Check for branch mention
        branch = self._extract_branch(text_clean)
        if branch:
            return f"BRANCH:{branch}"

        # 3. Intent matching
        scores = {}
        tokens = set(re.findall(r'\w+', text_clean))
        for intent, keywords in self.INTENT_MAP.items():
            score = 0
            for kw in keywords:
                if kw in text_clean:
                    score += (2 if len(kw.split()) > 1 else 1)  # multi-word bonus
                elif kw in tokens:
                    score += 1
            # fuzzy match on tokens
            for token in tokens:
                if len(token) >= 4:
                    for kw in keywords:
                        if difflib.SequenceMatcher(None, token, kw).ratio() > 0.82:
                            score += 0.5
            scores[intent] = score

        if scores:
            best = max(scores, key=scores.get)
            if scores[best] > 0:
                return best

        return "UNKNOWN"

    def extract_branch_from_text(self, text: str):
        """Return best-matching branch name or None"""
        return self._extract_branch(text.strip().lower())

    # ── Helpers ──
    def _is_follow_up(self, text: str) -> bool:
        for pat in self.FOLLOW_UP_PATTERNS:
            if re.match(pat, text):
                return True
        return False

    def _extract_branch(self, text: str):
        # Exact match
        for name in self.branch_names:
            if name.lower() in text:
                return name
        # Fuzzy
        best_match = difflib.get_close_matches(text, self.branch_names_lower, n=1, cutoff=0.6)
        if best_match:
            idx = self.branch_names_lower.index(best_match[0])
            return self.branch_names[idx]
        # Token-level partial
        tokens = text.split()
        for token in tokens:
            if len(token) >= 4:
                matches = difflib.get_close_matches(token, self.branch_names_lower, n=1, cutoff=0.7)
                if matches:
                    idx = self.branch_names_lower.index(matches[0])
                    return self.branch_names[idx]
        return None


# ─────────────────────────────────────────────
# AI ASSISTANT  (with SmartIntentEngine)
# ─────────────────────────────────────────────
class AIAssistant:
    def __init__(self, df):
        self.df = df
        self.conversation_history = []
        self.intent_engine = SmartIntentEngine(df)
        # Context memory
        self.last_intent = None
        self.last_branch = None
        self.last_response = ""

        api_key = os.getenv('OPENAI_API_KEY')
        if api_key:
            try:
                from openai import OpenAI
                self.client = OpenAI(api_key=api_key)
                self.api_available = True
            except ImportError:
                self.client = None
                self.api_available = False
        else:
            self.client = None
            self.api_available = False

    # ── Main entry ──
    def chat(self, user_query: str) -> str:
        """Process user query and return response"""
        try:
            intent = self.intent_engine.detect_intent(user_query)

        # Resolve follow-ups using context
            if intent == "FOLLOW_UP":
                intent = self.last_intent or "overview"

        # Branch mention → store & adjust intent
            if intent.startswith("BRANCH:"):
                branch_name = intent.split(":", 1)[1]
                self.last_branch = branch_name
                intent = "branch_detail"
            else:
            # Check if there's a branch embedded even if intent is something else
                found_branch = self.intent_engine.extract_branch_from_text(user_query)
                if found_branch:
                    self.last_branch = found_branch

        # Try GPT-4o first
            if self.api_available and self.client:
                try:
                    response = self._call_openai(user_query, intent)
                    self._store(user_query, response, intent)
                    return response
                except Exception as e:
                # If OpenAI fails, fall back to local
                    print(f"OpenAI API error: {e}")
                    pass

        # Fallback (local)
            response = self._local_response(intent, user_query)
            self._store(user_query, response, intent)
            return response
        
        except Exception as e:
            error_msg = f"I encountered an issue: {str(e)}. Let me try to help you anyway!"
            return error_msg

    def _store(self, query, response, intent):
        self.conversation_history.append({'user': query, 'assistant': response, 'timestamp': datetime.now()})
        self.last_intent = intent
        self.last_response = response

    # ── OpenAI call ──
    def _call_openai(self, user_query, resolved_intent):
        data_context = self._prepare_data_context()
        system = f"""You are BankVista AI, a supportive banking performance coach.

COMMUNICATION RULES:
✓ Start with strengths. Provide context. Use "you might consider", "one option is".
✓ Acknowledge constraints. Frame as opportunities. Offer 2-3 options. End with encouragement.
✗ NEVER use: critical, severe, urgent, failure, worst, must, required.
✗ No rankings or judgmental comparisons. No commands.

IMPORTANT – You understand casual / layman language:
- "bad loans" = NPA (Non-Performing Assets)
- "cheap deposits" = CASA (Current Account Savings Account)
- "stuck loans" = NPA
- "how far from target" = target gap analysis
- "weird numbers" = anomaly detection
- "who is doing well" = top performers

The user's detected intent is: {resolved_intent}
Last referenced branch (if any): {self.last_branch}

CURRENT DATA:
{data_context}

Respond naturally and supportively. If the user asked something vague, give a helpful overview and ask a gentle follow-up."""

        messages = [{"role": "system", "content": system}]
        for msg in self.conversation_history[-4:]:
            messages.append({"role": "user", "content": msg['user']})
            messages.append({"role": "assistant", "content": msg['assistant']})
        messages.append({"role": "user", "content": user_query})

        resp = self.client.chat.completions.create(model="gpt-4o", messages=messages, max_tokens=1500, temperature=0.7)
        return resp.choices[0].message.content

    # ── Local fallback responses (intent-driven) ──
    def _local_response(self, intent, raw_query):
        df = self.df

        if intent == "npa":
            return self._resp_npa()
        elif intent == "casa":
            return self._resp_casa()
        elif intent == "compare" or intent == "top_performers":
            return self._resp_compare()
        elif intent == "weak_branches":
            return self._resp_weak()
        elif intent == "deposits":
            return self._resp_deposits()
        elif intent == "advances":
            return self._resp_advances()
        elif intent == "staff":
            return self._resp_staff()
        elif intent == "target":
            return self._resp_target()
        elif intent == "anomaly":
            return self._resp_anomaly()
        elif intent == "zone":
            return self._resp_zone()
        elif intent == "branch_detail":
            return self._resp_branch(self.last_branch)
        elif intent == "overview":
            return self._resp_overview()
        else:
            return self._resp_help()

    # ── Individual response builders ──
    def _resp_npa(self):
        high = self.df.nlargest(3, 'NPA_Percent')
        avg = self.df['NPA_Percent'].mean()
        r = f"📊 **NPA (Bad Loan) Overview**\n\nOrganization average: **{avg:.2f}%**\n\n"
        r += "Here are the branches where a bit of extra attention could make a big difference:\n\n"
        for _, row in high.iterrows():
            tag = "🔴 needs attention" if row['NPA_Percent'] > 6 else "🟡 slightly elevated"
            r += f"• **{row['Branch_Name']}** — {row['NPA_Percent']:.2f}% {tag}\n"
        r += "\n💡 **Some ideas:** Early outreach to borrowers, understanding local challenges, and focused recovery efforts can help. A 1–2% drop in 60–90 days would be a great win!\n"
        r += "\n*Feel free to ask about a specific branch for a deeper look!*"
        return r

    def _resp_casa(self):
        low = self.df.nsmallest(3, 'CASA_Percent')
        avg = self.df['CASA_Percent'].mean()
        r = f"💰 **CASA (Cheap Deposit) Opportunities**\n\nOrg average: **{avg:.1f}%**\n\n"
        r += "These branches have the most room to grow their current & savings accounts:\n\n"
        for _, row in low.iterrows():
            r += f"• **{row['Branch_Name']}** — {row['CASA_Percent']:.1f}%\n"
        r += "\n💡 **Ideas to explore:** Partnering with local colleges, targeting small business owners for current accounts, and running savings campaigns can all help move the needle.\n"
        return r

    def _resp_compare(self):
        df = self.df.copy()
        df['dep_pct'] = (df['Total_Deposits'] / df['Deposit_Target'] * 100).round(1)
        df['adv_pct'] = (df['Advances'] / df['Advance_Target'] * 100).round(1)
        df['score'] = (
            df['dep_pct'].clip(upper=100) * 0.3 +
            df['adv_pct'].clip(upper=100) * 0.3 +
            (100 - df['NPA_Percent'] * 10).clip(lower=0) * 0.25 +
            df['CASA_Percent'] * 0.15
        )
        df = df.sort_values('score', ascending=False)
        r = "🏆 **Branch Performance Ranking**\n\n"
        for i, (_, row) in enumerate(df.iterrows(), 1):
            emoji = "🥇" if i == 1 else "🥈" if i == 2 else "🥉" if i == 3 else f"#{i}"
            r += f"{emoji} **{row['Branch_Name']}** — Score: {row['score']:.1f} | Dep: {row['dep_pct']}% | Adv: {row['adv_pct']}%\n"
        r += "\n*Every branch brings unique strengths to the table!*"
        return r

    def _resp_weak(self):
        df = self.df.copy()
        df['dep_pct'] = (df['Total_Deposits'] / df['Deposit_Target'] * 100)
        df['adv_pct'] = (df['Advances'] / df['Advance_Target'] * 100)
        weak = df[(df['dep_pct'] < 90) | (df['adv_pct'] < 90) | (df['NPA_Percent'] > 5)]
        r = "🛟 **Branches That Could Use Some Support**\n\n"
        if weak.empty:
            return r + "Great news — all branches are performing solidly! 🎉"
        for _, row in weak.iterrows():
            issues = []
            if row['dep_pct'] < 90: issues.append(f"Deposits at {row['dep_pct']:.0f}%")
            if row['adv_pct'] < 90: issues.append(f"Advances at {row['adv_pct']:.0f}%")
            if row['NPA_Percent'] > 5: issues.append(f"NPA at {row['NPA_Percent']:.1f}%")
            r += f"• **{row['Branch_Name']}** — {', '.join(issues)}\n"
        r += "\n💡 These branches might benefit from targeted support plans. Want details on any specific one?"
        return r

    def _resp_deposits(self):
        df = self.df.copy()
        df['pct'] = (df['Total_Deposits'] / df['Deposit_Target'] * 100).round(1)
        r = f"🏦 **Deposit Performance**\n\nTotal deposits across all branches: **₹{df['Total_Deposits'].sum():.1f} Cr**\n\n"
        for _, row in df.sort_values('pct', ascending=False).iterrows():
            icon = "✅" if row['pct'] >= 100 else "⚠️" if row['pct'] >= 85 else "🔶"
            r += f"{icon} **{row['Branch_Name']}** — ₹{row['Total_Deposits']:.1f} Cr / ₹{row['Deposit_Target']:.1f} Cr ({row['pct']}%)\n"
        return r

    def _resp_advances(self):
        df = self.df.copy()
        df['pct'] = (df['Advances'] / df['Advance_Target'] * 100).round(1)
        r = f"📋 **Advances (Loans Given) Performance**\n\nTotal advances: **₹{df['Advances'].sum():.1f} Cr**\n\n"
        for _, row in df.sort_values('pct', ascending=False).iterrows():
            icon = "✅" if row['pct'] >= 100 else "⚠️" if row['pct'] >= 85 else "🔶"
            r += f"{icon} **{row['Branch_Name']}** — ₹{row['Advances']:.1f} Cr / ₹{row['Advance_Target']:.1f} Cr ({row['pct']}%)\n"
        return r

    def _resp_staff(self):
        df = self.df
        r = "👥 **Staff & Efficiency Overview**\n\n"
        r += f"Total staff across branches: **{df['Staff_Count'].sum()}**\n\n"
        for _, row in df.sort_values('Business_Per_Staff', ascending=False).iterrows():
            r += f"• **{row['Branch_Name']}** — {row['Staff_Count']} staff | Business/Staff: ₹{row['Business_Per_Staff']:.1f} Cr | Profit/Staff: ₹{row['Profit_Per_Staff']:.1f} Cr\n"
        r += "\n💡 Higher business-per-staff often comes from experience and local knowledge!"
        return r

    def _resp_target(self):
        df = self.df.copy()
        df['dep_pct'] = (df['Total_Deposits'] / df['Deposit_Target'] * 100).round(1)
        df['adv_pct'] = (df['Advances'] / df['Advance_Target'] * 100).round(1)
        r = "🎯 **Target Achievement Tracker**\n\n"
        for _, row in df.iterrows():
            r += f"**{row['Branch_Name']}**\n"
            r += f"  Deposits: {row['dep_pct']}% {'✅' if row['dep_pct'] >= 100 else '⚠️'} | "
            r += f"Advances: {row['adv_pct']}% {'✅' if row['adv_pct'] >= 100 else '⚠️'}\n"
        return r

    def _resp_anomaly(self):
        df = self.df.copy()
        anomalies = []
        for col in ['NPA_Percent', 'CASA_Percent', 'CD_Ratio', 'Profit_Per_Staff']:
            mean, std = df[col].mean(), df[col].std()
            if std == 0: continue
            for _, row in df.iterrows():
                z = abs(row[col] - mean) / std
                if z > 1.8:
                    anomalies.append({'branch': row['Branch_Name'], 'metric': col, 'value': row[col], 'z': z})
        r = "🔍 **Anomaly Detection Report**\n\n"
        if not anomalies:
            return r + "All metrics look normal across branches — no unusual patterns detected! 👍"
        for a in anomalies:
            direction = "high" if a['value'] > df[a['metric']].mean() else "low"
            r += f"• **{a['branch']}** — {a['metric']} is unusually **{direction}** at {a['value']:.2f} (Z-score: {a['z']:.2f})\n"
        r += "\n💡 These aren't necessarily bad — just worth a quick look to understand the context."
        return r

    def _resp_zone(self):
        df = self.df
        r = "🗺️ **Zone-wise Summary**\n\n"
        for zone in df['Zone'].unique():
            zdf = df[df['Zone'] == zone]
            r += f"### {zone}\n"
            r += f"  Branches: {len(zdf)} | Total Deposits: ₹{zdf['Total_Deposits'].sum():.1f} Cr | Avg NPA: {zdf['NPA_Percent'].mean():.2f}% | Avg CASA: {zdf['CASA_Percent'].mean():.1f}%\n\n"
        return r

    def _resp_branch(self, branch_name):
        if not branch_name:
            return "Could you tell me which branch you're asking about? I can help with any of them! 😊"
        row = self.df[self.df['Branch_Name'] == branch_name]
        if row.empty:
            return f"I couldn't find a branch matching '{branch_name}'. Could you clarify? Here are the available branches: {', '.join(self.df['Branch_Name'].tolist())}"
        row = row.iloc[0]
        dep_pct = row['Total_Deposits'] / row['Deposit_Target'] * 100
        adv_pct = row['Advances'] / row['Advance_Target'] * 100
        r = f"📍 **{branch_name} — Full Snapshot**\n\n"
        r += f"| Metric | Value |\n|---|---|\n"
        r += f"| Zone | {row['Zone']} |\n"
        r += f"| Deposits | ₹{row['Total_Deposits']:.1f} Cr ({dep_pct:.1f}% of target) |\n"
        r += f"| Advances | ₹{row['Advances']:.1f} Cr ({adv_pct:.1f}% of target) |\n"
        r += f"| NPA | {row['NPA_Percent']:.2f}% |\n"
        r += f"| CASA | {row['CASA_Percent']:.1f}% |\n"
        r += f"| Staff | {row['Staff_Count']} |\n"
        r += f"| Business/Staff | ₹{row['Business_Per_Staff']:.1f} Cr |\n"
        r += f"| Profit/Staff | ₹{row['Profit_Per_Staff']:.1f} Cr |\n"
        r += f"| CD Ratio | {row['CD_Ratio']:.1f}% |\n"
        return r

    def _resp_overview(self):
        df = self.df
        total_dep = df['Total_Deposits'].sum()
        total_adv = df['Advances'].sum()
        r = f"📊 **BankVista — Overall Health Check**\n\n"
        r += f"• **{len(df)} branches** across {df['Zone'].nunique()} zone(s)\n"
        r += f"• Total Deposits: **₹{total_dep:.1f} Cr** | Total Advances: **₹{total_adv:.1f} Cr**\n"
        r += f"• Avg NPA: **{df['NPA_Percent'].mean():.2f}%** | Avg CASA: **{df['CASA_Percent'].mean():.1f}%**\n"
        r += f"• Total Staff: **{df['Staff_Count'].sum()}**\n\n"
        r += "💬 **You can ask me things like:**\n"
        r += "  • \"Which branches have bad loan problems?\"\n"
        r += "  • \"How are deposits looking?\"\n"
        r += "  • \"Anything weird in the numbers?\"\n"
        r += "  • \"Tell me about Hyderabad Main\"\n"
        return r

    def _resp_help(self):
        r = "👋 **Hi! Here's what I understand:**\n\n"
        r += "🔴 **Bad loans / stuck loans / NPA** → I'll show you NPA analysis\n"
        r += "💰 **Savings / cheap deposits / CASA** → CASA opportunities\n"
        r += "🏆 **Who's doing well / best branch** → Top performers\n"
        r += "🛟 **Struggling / need help / behind** → Branches needing support\n"
        r += "🎯 **Target / gap / on track** → Target achievement\n"
        r += "🔍 **Anything weird / odd / anomaly** → Anomaly detection\n"
        r += "📍 **[Branch name]** → Full branch snapshot\n"
        r += "🗺️ **Zone / region** → Zone-wise summary\n\n"
        r += "Just type naturally — I'll figure out what you need! 😊"
        return r

    def _prepare_data_context(self):
        lines = [f"Total Branches: {len(self.df)}", f"Total Deposits: ₹{self.df['Total_Deposits'].sum():.2f} Cr",
                 f"Avg NPA: {self.df['NPA_Percent'].mean():.2f}%", f"Avg CASA: {self.df['CASA_Percent'].mean():.2f}%\n"]
        for _, row in self.df.iterrows():
            lines.append(f"{row['Branch_Name']}: Deposits ₹{row['Total_Deposits']:.1f}Cr, NPA {row['NPA_Percent']:.1f}%, CASA {row['CASA_Percent']:.1f}%, Staff {row['Staff_Count']}")
        return "\n".join(lines)

    def export_conversation(self):
        return json.dumps([{'user': m['user'], 'assistant': m['assistant'], 'timestamp': m['timestamp'].isoformat()} for m in self.conversation_history], indent=2)


# ─────────────────────────────────────────────
# PREDICTIVE ANALYTICS
# ─────────────────────────────────────────────
class PredictiveAnalytics:
    def __init__(self, df):
        self.df = df

    def predict_npa_trend(self, branch_name, months=6):
        branch = self.df[self.df['Branch_Name'] == branch_name].iloc[0]
        current_npa = branch['NPA_Percent']
        predictions = []
        npa = current_npa
        for m in range(1, months + 1):
            seasonal = 1.0 + (0.08 if m in [3, 6, 9, 12] else 0)
            noise = np.random.normal(0, 0.18)
            npa = max(0, npa * seasonal + noise)
            predictions.append({'month': m, 'predicted_npa': round(npa, 2), 'confidence': 'high' if m <= 2 else 'medium'})
        final = predictions[-1]['predicted_npa']
        risk = 'needs attention' if final > 6 else 'watch' if final > 3 else 'stable'
        return {'branch': branch_name, 'current_npa': branch['NPA_Percent'], 'predictions': predictions, 'risk_level': risk}

    def predict_target_achievement(self, branch_name):
        branch = self.df[self.df['Branch_Name'] == branch_name].iloc[0]
        dep_ach = (branch['Total_Deposits'] / branch['Deposit_Target']) * 100
        adv_ach = (branch['Advances'] / branch['Advance_Target']) * 100

        def prob(pct):
            if pct >= 95: return 95
            elif pct >= 85: return 70 + (pct - 85) * 2.5
            else: return max(30, 50 + (pct - 75) * 2)

        dp, ap = min(100, prob(dep_ach)), min(100, prob(adv_ach))
        return {'branch': branch_name, 'deposit_probability': dp, 'advance_probability': ap,
                'overall_probability': (dp + ap) / 2, 'recommendation': 'On track ✅' if dp > 80 else 'Needs attention ⚠️'}


# ─────────────────────────────────────────────
# ENHANCED BRANCH ANALYZER
# ─────────────────────────────────────────────
class EnhancedBankVista:
    def __init__(self, data):
        self.data = data
        self.insights = {'needs_support': [], 'watch_area': [], 'strengths': []}

    def analyze(self):
        self._check_deposits()
        self._check_advances()
        self._check_npa()
        self._check_casa()
        grade, score = self._calc_grade()
        return {'grade': grade, 'score': score, 'insights': self.insights}

    def _check_deposits(self):
        t, tgt = float(self.data.get('Total_Deposits', 0)), float(self.data.get('Deposit_Target', 1))
        pct = t / tgt * 100 if tgt else 0
        bucket = 'needs_support' if pct < 85 else 'watch_area' if pct < 95 else 'strengths'
        label = 'Growth Opportunity' if pct < 85 else 'Nearly There' if pct < 95 else 'Strong Performance'
        self.insights[bucket].append({'title': f'Deposits - {label}', 'detail': f'{pct:.1f}% achieved'})

    def _check_advances(self):
        t, tgt = float(self.data.get('Advances', 0)), float(self.data.get('Advance_Target', 1))
        pct = t / tgt * 100 if tgt else 0
        bucket = 'needs_support' if pct < 85 else 'watch_area' if pct < 95 else 'strengths'
        label = 'Room for Growth' if pct < 85 else 'Good Progress' if pct < 95 else 'Excellent'
        self.insights[bucket].append({'title': f'Advances - {label}', 'detail': f'{pct:.1f}% achieved'})

    def _check_npa(self):
        npa = float(self.data.get('NPA_Percent', 0))
        if npa > 6: self.insights['needs_support'].append({'title': 'NPA - Needs Focused Attention', 'detail': f'{npa:.2f}% (Target: 3%)'})
        elif npa > 3: self.insights['watch_area'].append({'title': 'NPA - Monitor Closely', 'detail': f'{npa:.2f}%'})
        else: self.insights['strengths'].append({'title': 'NPA - Healthy', 'detail': f'{npa:.2f}%'})

    def _check_casa(self):
        casa = float(self.data.get('CASA_Percent', 0))
        if casa < 30: self.insights['watch_area'].append({'title': 'CASA - Development Opportunity', 'detail': f'{casa:.1f}%'})
        elif casa >= 40: self.insights['strengths'].append({'title': 'CASA - Excellent', 'detail': f'{casa:.1f}%'})

    def _calc_grade(self):
        try:
            d_a, d_t = float(self.data.get('Total_Deposits', 0)), float(self.data.get('Deposit_Target', 1))
            a_a, a_t = float(self.data.get('Advances', 0)), float(self.data.get('Advance_Target', 1))
            npa, casa = float(self.data.get('NPA_Percent', 0)), float(self.data.get('CASA_Percent', 0))
            s = (min(d_a / d_t * 25, 25) if d_t else 0) + (min(a_a / a_t * 25, 25) if a_t else 0)
            s += 20 if npa <= 3 else (12 if npa <= 6 else 5)
            s += 10 if casa >= 40 else (5 if casa >= 30 else 2)
            s = round(s, 1)
            g = "A+" if s >= 70 else "A" if s >= 60 else "B" if s >= 50 else "C" if s >= 40 else "D"
            return g, s
        except:
            return "N/A", 0


# ─────────────────────────────────────────────
# ANOMALY DETECTOR
# ─────────────────────────────────────────────
class AnomalyDetector:
    def __init__(self, df):
        self.df = df

    def detect(self):
        anomalies = []
        metrics = ['NPA_Percent', 'CASA_Percent', 'CD_Ratio', 'Profit_Per_Staff', 'Business_Per_Staff']
        for col in metrics:
            if col not in self.df.columns: continue
            mean, std = self.df[col].mean(), self.df[col].std()
            if std == 0: continue
            for _, row in self.df.iterrows():
                z = (row[col] - mean) / std
                if abs(z) > 1.7:
                    anomalies.append({
                        'branch': row['Branch_Name'], 'metric': col,
                        'value': round(row[col], 2), 'z_score': round(z, 2),
                        'direction': 'HIGH' if z > 0 else 'LOW', 'mean': round(mean, 2)
                    })
        return anomalies


# ─────────────────────────────────────────────
# EXCEL EXPORT
# ─────────────────────────────────────────────
def create_dynamic_excel(df):
    output = io.BytesIO()
    wb = Workbook()
    wb.remove(wb.active)
    thin_border = Border(left=Side('thin'), right=Side('thin'), top=Side('thin'), bottom=Side('thin'))

    # Hidden data sheet
    ws_data = wb.create_sheet("_Data")
    ws_data.sheet_state = 'hidden'
    for ci, h in enumerate(df.columns.tolist(), 1):
        ws_data.cell(1, ci, h).font = Font(bold=True)
    for ri, row in enumerate(df.itertuples(index=False), 2):
        for ci, v in enumerate(row, 1):
            ws_data.cell(ri, ci, v)

    # Dashboard
    ws = wb.create_sheet("Dashboard", 0)
    ws.merge_cells('A1:H1')
    ws['A1'] = "🤖 BANKVISTA AI – DYNAMIC DASHBOARD"
    ws['A1'].font = Font(bold=True, size=16, color="FFFFFF")
    ws['A1'].fill = PatternFill("solid", fgColor="667EEA")
    ws['A1'].alignment = Alignment(horizontal='center', vertical='center')
    ws.row_dimensions[1].height = 30

    ws['A3'] = "Select Branch:"
    ws['A3'].font = Font(bold=True, size=12)
    ws.merge_cells('B3:D3')
    ws['B3'] = df.iloc[0]['Branch_Name']
    ws['B3'].font = Font(size=12, bold=True)
    ws['B3'].fill = PatternFill("solid", fgColor="FFF2CC")

    last_row = len(df) + 1
    dv = DataValidation(type="list", formula1=f"=_Data!$B$2:$B${last_row}", allow_blank=False)
    dv.add('B3')
    ws.add_data_validation(dv)

    ws['F3'] = "Date:"
    ws['G3'] = date.today().strftime('%d-%b-%Y')

    # Info
    ws['A5'] = "Branch ID:"
    ws['B5'] = '=INDEX(_Data!$A:$A,MATCH(B3,_Data!$B:$B,0))'
    ws['A6'] = "Zone:"
    ws['B6'] = '=INDEX(_Data!$C:$C,MATCH(B3,_Data!$B:$B,0))'

    # Grade / Score
    # ── FIXED: Grade / Score Headers & Values ──
    # GRADE Header - set value BEFORE merging
    ws['A8'] = "GRADE"
    ws['A8'].font = Font(bold=True, color="FFFFFF")
    ws['A8'].fill = PatternFill("solid", fgColor="27AE60")
    ws['A8'].alignment = Alignment(horizontal='center')
    ws.merge_cells('A8:B8')

    # GRADE Value - set value BEFORE merging
    ws['A9'] = '=IF(C9>=70,"A+",IF(C9>=60,"A",IF(C9>=50,"B",IF(C9>=40,"C","D"))))'
    ws['A9'].font = Font(bold=True, size=16)
    ws['A9'].alignment = Alignment(horizontal='center')
    ws.merge_cells('A9:B9')

    # SCORE Header - set value BEFORE merging
    ws['D8'] = "SCORE"
    ws['D8'].font = Font(bold=True, color="FFFFFF")
    ws['D8'].fill = PatternFill("solid", fgColor="3498DB")
    ws['D8'].alignment = Alignment(horizontal='center')
    ws.merge_cells('D8:E8')

    # SCORE Value - set value BEFORE merging
    ws['D9'] = '=ROUND(C9,0)&"/80"'
    ws['D9'].font = Font(bold=True, size=14)
    ws['D9'].alignment = Alignment(horizontal='center')
    ws.merge_cells('D9:E9')

    # Hidden score calc (Column C9)
    ws['C9'] = ('=MIN(INDEX(_Data!$D:$D,MATCH(B3,_Data!$B:$B,0))/INDEX(_Data!$E:$E,MATCH(B3,_Data!$B:$B,0))*25,25)'
                '+MIN(INDEX(_Data!$F:$F,MATCH(B3,_Data!$B:$B,0))/INDEX(_Data!$G:$G,MATCH(B3,_Data!$B:$B,0))*25,25)'
                '+IF(INDEX(_Data!$H:$H,MATCH(B3,_Data!$B:$B,0))<=3,20,IF(INDEX(_Data!$H:$H,MATCH(B3,_Data!$B:$B,0))<=6,12,5))'
                '+IF(INDEX(_Data!$J:$J,MATCH(B3,_Data!$B:$B,0))>=40,10,5)')
    ws.column_dimensions['C'].hidden = True

    # Metrics table
    ws.merge_cells('A11:H11'); ws['A11'] = "KEY METRICS"
    ws['A11'].font = Font(bold=True, size=12, color="FFFFFF")
    ws['A11'].fill = PatternFill("solid", fgColor="4472C4")
    ws['A11'].alignment = Alignment(horizontal='center')

    for i, h in enumerate(['Metric', 'Actual', 'Target', 'Gap', 'Achievement %', 'Status'], 1):
        c = ws.cell(12, i); c.value = h; c.font = Font(bold=True, color="FFFFFF")
        c.fill = PatternFill("solid", fgColor="5B9BD5"); c.alignment = Alignment(horizontal='center'); c.border = thin_border

    metrics = [("Deposits (Cr)", 4, 5), ("Advances (Cr)", 6, 7), ("NPA %", 8, None), ("CASA %", 10, None)]
    row = 13
    for metric, ac, tc in metrics:
        ws.cell(row, 1, metric)
        acl = get_column_letter(ac)
        ws.cell(row, 2, f'=INDEX(_Data!${acl}:${acl},MATCH(B3,_Data!$B:$B,0))')
        ws.cell(row, 2).number_format = '#,##0.00'
        if tc:
            tcl = get_column_letter(tc)
            ws.cell(row, 3, f'=INDEX(_Data!${tcl}:${tcl},MATCH(B3,_Data!$B:$B,0))')
            ws.cell(row, 3).number_format = '#,##0.00'
            ws.cell(row, 4, f'=B{row}-C{row}'); ws.cell(row, 4).number_format = '#,##0.00'
            ws.cell(row, 5, f'=B{row}/C{row}'); ws.cell(row, 5).number_format = '0.0%'
            ws.cell(row, 6, f'=IF(B{row}>=C{row},"✅ On Track","⚠️ Gap")')
        else:
            if metric == "NPA %":
                ws.cell(row, 3, "3.00%"); ws.cell(row, 6, f'=IF(B{row}<=3,"✅ Good","⚠️ High")')
            else:
                ws.cell(row, 3, "40.00%"); ws.cell(row, 6, f'=IF(B{row}>=40,"✅ Excellent","⚠️ Low")')
        for col in range(1, 7):
            ws.cell(row, col).border = thin_border; ws.cell(row, col).alignment = Alignment(horizontal='center')
        row += 1

    # Instructions
    ws.merge_cells('A18:H18'); ws['A18'] = "💡 INSTRUCTIONS"
    ws['A18'].font = Font(bold=True, size=12, color="FFFFFF")
    ws['A18'].fill = PatternFill("solid", fgColor="27AE60"); ws['A18'].alignment = Alignment(horizontal='center')
    for idx, inst in enumerate(["1. Click dropdown in B3 to select any branch",
                                 "2. All metrics update automatically via formulas",
                                 "3. Works completely offline – share freely",
                                 "4. Check 'All Branches' sheet for full summary"], 19):
        ws.merge_cells(f'A{idx}:H{idx}'); ws[f'A{idx}'] = inst

    for c in 'ABCDEFGH': ws.column_dimensions[c].width = 16

    # All Branches summary
    ws_s = wb.create_sheet("All Branches")
    ws_s.merge_cells('A1:G1'); ws_s['A1'] = "ALL BRANCHES SUMMARY"
    ws_s['A1'].font = Font(bold=True, size=14, color="FFFFFF")
    ws_s['A1'].fill = PatternFill("solid", fgColor="1F4E78"); ws_s['A1'].alignment = Alignment(horizontal='center')

    for i, h in enumerate(['Branch', 'Zone', 'Deposits %', 'Advances %', 'NPA %', 'CASA %', 'Grade'], 1):
        c = ws_s.cell(3, i); c.value = h; c.font = Font(bold=True, color="FFFFFF")
        c.fill = PatternFill("solid", fgColor="4472C4"); c.alignment = Alignment(horizontal='center'); c.border = thin_border

    for ri, row in df.iterrows():
        r = ri + 4
        dep_pct = row['Total_Deposits'] / row['Deposit_Target'] * 100
        adv_pct = row['Advances'] / row['Advance_Target'] * 100
        score = (min(dep_pct / 100 * 25, 25) + min(adv_pct / 100 * 25, 25) +
                 (20 if row['NPA_Percent'] <= 3 else 12 if row['NPA_Percent'] <= 6 else 5) +
                 (10 if row['CASA_Percent'] >= 40 else 5))
        grade = "A+" if score >= 70 else "A" if score >= 60 else "B" if score >= 50 else "C"
        vals = [row['Branch_Name'], row['Zone'], f"{dep_pct:.1f}%", f"{adv_pct:.1f}%",
                f"{row['NPA_Percent']:.2f}%", f"{row['CASA_Percent']:.1f}%", grade]
        for ci, v in enumerate(vals, 1):
            ws_s.cell(r, ci, v).border = thin_border
            ws_s.cell(r, ci).alignment = Alignment(horizontal='center')

    for i in range(1, 8): ws_s.column_dimensions[get_column_letter(i)].width = 18
    wb.save(output); output.seek(0)
    return output


# ─────────────────────────────────────────────
# SAMPLE DATA
# ─────────────────────────────────────────────
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
        if ext == 'csv': return pd.read_csv(file)
        elif ext in ['xlsx','xls']: return pd.read_excel(file)
        else: st.error("Unsupported format"); return None
    except Exception as e: st.error(f"Error: {e}"); return None


# ─────────────────────────────────────────────
# MAIN APP
# ─────────────────────────────────────────────
def main():
    # ── Hero ──
    st.markdown("""
    <div class="hero-container">
        <h1 class="main-title">🤖 BankVista AI</h1>
        <p class="subtitle">Next-Gen AI Banking Analytics · Understands Plain English</p>
        <div class="feature-pills">
            <div class="feature-pill"><span>💬</span> Smart Chat</div>
            <div class="feature-pill"><span>📈</span> Predictions</div>
            <div class="feature-pill"><span>🎯</span> Risk Scoring</div>
            <div class="feature-pill"><span>🔍</span> Anomaly Detection</div>
            <div class="feature-pill"><span>📊</span> Branch Compare</div>
            <div class="feature-pill"><span>🗺️</span> Zone Analytics</div>
            <div class="feature-pill"><span>👥</span> Staff Efficiency</div>
            <div class="feature-pill"><span>📥</span> Excel Export</div>
        </div>
    </div>
    """, unsafe_allow_html=True)

    # ── Sidebar ──
    with st.sidebar:
        st.markdown("### 📤 Upload Data")
        uploaded = st.file_uploader("Choose file (CSV / Excel)", type=['csv','xlsx','xls'])
        if st.button("📊 Try Sample Data", key="sidebar_sample_data_btn"):  # ← Added unique key
            st.session_state['use_sample'] = True
            st.session_state['uploaded_df'] = sample_data()
            if 'ai_assistant' in st.session_state:
                del st.session_state['ai_assistant']
                st.session_state['chat_messages'] = []
            st.rerun()  # ← Added explicit rerun

        st.markdown("---")
        st.markdown("""
        <div style="background:#d1fae5;padding:1.2rem;border-radius:12px;border-left:4px solid #10b981;">
            <h4 style="margin-top:0;color:#065f46;">🤖 AI Understands</h4>
            <p style="margin-bottom:0.4rem;color:#065f46;">✅ "Which loans are bad?" → NPA</p>
            <p style="margin-bottom:0.4rem;color:#065f46;">✅ "Cheap deposits?" → CASA</p>
            <p style="margin-bottom:0.4rem;color:#065f46;">✅ "Anything weird?" → Anomalies</p>
            <p style="margin-bottom:0.4rem;color:#065f46;">✅ "Tell me about Warangal"</p>
            <p style="margin-bottom:0;color:#065f46;">✅ Follow-ups like "tell me more"</p>
        </div>
        """, unsafe_allow_html=True)

        st.markdown("---")
        st.markdown("""
        <div style="background:#dbeafe;padding:1.2rem;border-radius:12px;border-left:4px solid #3b82f6;">
            <h4 style="margin-top:0;color:#1e40af;">📊 Features</h4>
            <p style="margin-bottom:0.3rem;color:#1e40af;">• AI Chat (plain language)</p>
            <p style="margin-bottom:0.3rem;color:#1e40af;">• Branch Comparison</p>
            <p style="margin-bottom:0.3rem;color:#1e40af;">• Heatmap Visualization</p>
            <p style="margin-bottom:0.3rem;color:#1e40af;">• Anomaly Detection</p>
            <p style="margin-bottom:0.3rem;color:#1e40af;">• Staff Efficiency</p>
            <p style="margin-bottom:0.3rem;color:#1e40af;">• Zone Analytics</p>
            <p style="margin-bottom:0;color:#1e40af;">• Dynamic Excel Export</p>
        </div>
        """, unsafe_allow_html=True)

    # ── Load data ──
    if uploaded:
        df = load_file(uploaded)
        if df is not None:
            st.session_state['uploaded_df'] = df
            st.session_state['use_sample'] = False
            if 'ai_assistant' in st.session_state:
                del st.session_state['ai_assistant']
                st.session_state['chat_messages'] = []

    if st.session_state.get('use_sample') or 'uploaded_df' in st.session_state:
        df = st.session_state['uploaded_df']

        # Init AI
        if 'ai_assistant' not in st.session_state:
            st.session_state['ai_assistant'] = AIAssistant(df)
            st.session_state['chat_messages'] = []

        predictive = PredictiveAnalytics(df)

        # ── Status bar ──
        st.markdown(f"""
        <div style="background:#d1fae5;padding:1rem 1.5rem;border-radius:12px;margin-bottom:1.5rem;border:1px solid #10b981;">
            <p style="text-align:center;font-size:1rem;margin:0;color:#065f46;">
                ✅ <strong>{len(df)} branches loaded</strong> &nbsp;|&nbsp; 🤖 <strong>AI Ready</strong> &nbsp;|&nbsp;
                💰 <strong>₹{df['Total_Deposits'].sum():.1f} Cr deposits</strong> &nbsp;|&nbsp;
                👥 <strong>{df['Staff_Count'].sum()} staff</strong>
            </p>
        </div>
        """, unsafe_allow_html=True)

        # ── Summary KPIs ──
        c1, c2, c3, c4 = st.columns(4)
        avg_npa = df['NPA_Percent'].mean()
        avg_casa = df['CASA_Percent'].mean()
        dep_ach = (df['Total_Deposits'].sum() / df['Deposit_Target'].sum() * 100)

        for col, val, label, trend in [
            (c1, f"{dep_ach:.0f}%", "Deposit Achievement", "↗ vs target"),
            (c2, f"{avg_npa:.1f}%", "Avg NPA", "Monitored daily"),
            (c3, f"{avg_casa:.1f}%", "Avg CASA", "Growing"),
            (c4, f"₹{df['Business_Per_Staff'].mean():.0f}Cr", "Avg Biz/Staff", "Productivity"),
        ]:
            with col:
                st.markdown(f"""
                <div class="stat-card">
                    <div class="stat-value">{val}</div>
                    <div class="stat-label">{label}</div>
                    <div class="stat-trend">{trend}</div>
                </div>
                """, unsafe_allow_html=True)

        # ── TABS ──
        tab1, tab2, tab3, tab4, tab5, tab6, tab7 = st.tabs([
            "💬 AI Coach", "📈 Predictions", "📊 Dashboard",
            "🔄 Compare", "🗺️ Heatmaps", "🔍 Anomalies", "📥 Export"
        ])

        # ════════════════════════════════════════
        # TAB 1 — AI COACH
        # ════════════════════════════════════════
        with tab1:
            st.markdown('<p class="section-header">💬 Chat with Your Data — Plain English OK!</p>', unsafe_allow_html=True)

            # Chat display
            st.markdown('<div class="ai-chat-container">', unsafe_allow_html=True)
            if not st.session_state['chat_messages']:
                st.markdown('<div class="ai-message">🤖 Hi! I understand plain language. Try asking:<br>'
                            '• "Which loans are bad?"<br>'
                            '• "How are deposits looking?"<br>'
                            '• "Anything weird in the numbers?"<br>'
                            '• "Tell me about Hyderabad Main"<br><br>'
                            'Just type naturally — I\'ll figure it out! 😊</div>', unsafe_allow_html=True)
            for msg in st.session_state['chat_messages']:
                if msg['role'] == 'user':
                    st.markdown(f'<div class="user-message">👤 {msg["content"]}</div>', unsafe_allow_html=True)
                else:
                    st.markdown(f'<div class="ai-message">🤖 {msg["content"]}</div>', unsafe_allow_html=True)
            st.markdown('</div>', unsafe_allow_html=True)

            # Input row
            col1, col2 = st.columns([5, 1])
            with col1:
                user_input = st.text_input("Ask anything...", key="chat_input", label_visibility="collapsed",
                                           placeholder='e.g. "Which branches have bad loans?" or "Tell me about Warangal"')
            with col2:
                send_btn = st.button("Send 🚀", key="send_msg_btn")  # ← Added unique key

            if send_btn and user_input:
                with st.spinner("Thinking..."):  # ← Added spinner
                    st.session_state['chat_messages'].append({'role': 'user', 'content': user_input})
                    response = st.session_state['ai_assistant'].chat(user_input)
                    st.session_state['chat_messages'].append({'role': 'assistant', 'content': response})
                    st.rerun()

            # Quick buttons – expanded set
            st.markdown("**💡 Quick Questions:**")
            row1 = st.columns(4)
            quick_qs = [
                ("🔴 Bad Loans?", "Which branches have bad loans?"),
                ("💰 Cheap Deposits?", "Where can we get cheap deposits?"),
                ("🏆 Who's doing well?", "Which branches are doing great?"),
                ("🛟 Need help?", "Which branches are struggling?"),
            ]
            for i, (col, (label, q)) in enumerate(zip(row1, quick_qs)):
                with col:
                    if st.button(label, key=f"quick1_{i}"):  # ← Added unique key
                        with st.spinner("Thinking..."):  # ← Added spinner
                            st.session_state['chat_messages'].append({'role': 'user', 'content': q})
                            st.session_state['chat_messages'].append({'role': 'assistant', 'content': st.session_state['ai_assistant'].chat(q)})
                        st.rerun()  # ← Added explicit rerun

            row2 = st.columns(4)
            quick_qs2 = [
                ("🎯 On Track?", "Are we meeting our targets?"),
                ("🔍 Anything weird?", "Are there any anomalies or strange numbers?"),
                ("🗺️ Zone Summary", "Give me a zone-wise summary"),
                ("👥 Staff Stats", "How is staff efficiency?"),
            ]
            for i, (col, (label, q)) in enumerate(zip(row2, quick_qs2)):
                with col:
                    if st.button(label, key=f"quick2_{i}"):  # ← Added unique key
                        with st.spinner("Thinking..."):  # ← Added spinner
                            st.session_state['chat_messages'].append({'role': 'user', 'content': q})
                            st.session_state['chat_messages'].append({'role': 'assistant', 'content': st.session_state['ai_assistant'].chat(q)})
                        st.rerun()  # ← Added explicit rerun

        # ════════════════════════════════════════
        # TAB 2 — PREDICTIONS
        # ════════════════════════════════════════
        with tab2:
            st.markdown('<p class="section-header">📈 Predictive Analytics</p>', unsafe_allow_html=True)
            selected = st.selectbox("Select Branch:", df['Branch_Name'].tolist(), key="pred_branch")

            col1, col2 = st.columns(2)

            with col1:
                st.markdown("### 📉 NPA Forecast (6 Months)")
                npa_pred = predictive.predict_npa_trend(selected, months=6)
                color_map = {'stable': '#10b981', 'watch': '#f59e0b', 'needs attention': '#ef4444'}

                st.markdown(f"""
                <div class="metric-card">
                    <p style="font-size:1rem;margin-bottom:0.3rem;"><strong>Current NPA:</strong> {npa_pred['current_npa']:.2f}%</p>
                    <p style="font-size:1rem;margin:0;"><strong>Outlook:</strong>
                        <span style="color:{color_map.get(npa_pred['risk_level'],'#374151')};font-weight:700;">
                            {npa_pred['risk_level'].upper()}
                        </span>
                    </p>
                </div>
                """, unsafe_allow_html=True)

                months_x = [0] + [p['month'] for p in npa_pred['predictions']]
                npa_y = [npa_pred['current_npa']] + [p['predicted_npa'] for p in npa_pred['predictions']]

                fig = go.Figure()
                fig.add_trace(go.Scatter(x=months_x, y=npa_y, mode='lines+markers',
                              line=dict(color='#06b6d4', width=3), marker=dict(size=8, color='#06b6d4'),
                              name="Predicted NPA"))
                fig.add_hline(y=3, line_dash="dash", line_color="#10b981", annotation_text="Target: 3%")
                fig.add_hline(y=6, line_dash="dash", line_color="#f59e0b", annotation_text="Watch: 6%")
                fig.update_layout(title="NPA Trend Forecast", height=320, plot_bgcolor='white', paper_bgcolor='white',
                                  xaxis_title="Month", yaxis_title="NPA %")
                st.plotly_chart(fig, key="npa_forecast_chart")

            with col2:
                st.markdown("### 🎯 Target Achievement Probability")
                target_pred = predictive.predict_target_achievement(selected)
                bar_color = "#10b981" if target_pred['overall_probability'] > 75 else "#f59e0b"

                fig = go.Figure(go.Indicator(mode="gauge+number", value=target_pred['overall_probability'],
                    title={'text': "Success Probability"},
                    gauge={'axis': {'range': [0, 100]}, 'bar': {'color': bar_color},
                           'steps': [{'range': [0, 50], 'color': 'rgba(239,68,68,0.15)'},
                                     {'range': [50, 75], 'color': 'rgba(245,158,11,0.15)'},
                                     {'range': [75, 100], 'color': 'rgba(16,185,129,0.15)'}],
                           'threshold': {'line': {'color': '#1f2937', 'width': 3}, 'thickness': 0.7, 'value': 75}}))
                fig.update_layout(height=320, plot_bgcolor='white', paper_bgcolor='white')
                st.plotly_chart(fig, key="target_probability_gauge")

                st.markdown(f"""
                <div class="metric-card">
                    <p style="font-size:0.95rem;margin-bottom:0.3rem;"><strong>Deposit Prob:</strong> {target_pred['deposit_probability']:.0f}%</p>
                    <p style="font-size:0.95rem;margin-bottom:0.3rem;"><strong>Advance Prob:</strong> {target_pred['advance_probability']:.0f}%</p>
                    <p style="font-size:1.05rem;font-weight:bold;margin:0;"><strong>Status:</strong> {target_pred['recommendation']}</p>
                </div>
                """, unsafe_allow_html=True)

        # ════════════════════════════════════════
        # TAB 3 — DASHBOARD
        # ════════════════════════════════════════
        with tab3:
            st.markdown('<p class="section-header">📊 Branch Performance Dashboard</p>', unsafe_allow_html=True)
            selected = st.selectbox("Select Branch:", df['Branch_Name'].tolist(), key="dash_branch")
            branch_data = df[df['Branch_Name'] == selected].iloc[0].to_dict()
            analysis = EnhancedBankVista(branch_data).analyze()

            # KPI row
            grade_emoji = "🟢" if analysis['grade'] in ['A+','A'] else "🟡" if analysis['grade'] == 'B' else "🔴"
            kpis = [
                (f"{grade_emoji} {analysis['grade']}", "Grade"),
                (f"{analysis['score']}/80", "Score"),
                (f"{len(analysis['insights']['needs_support'])}", "Needs Support"),
                (f"{len(analysis['insights']['strengths'])}", "Strengths"),
            ]
            cols = st.columns(4)
            for col, (val, label) in zip(cols, kpis):
                with col:
                    st.markdown(f"""
                    <div class="metric-card" style="text-align:center;">
                        <div class="stat-value">{val}</div>
                        <div class="stat-label">{label}</div>
                    </div>
                    """, unsafe_allow_html=True)

            # Radar chart
            row = df[df['Branch_Name'] == selected].iloc[0]
            dep_pct = min(row['Total_Deposits'] / row['Deposit_Target'] * 100, 100)
            adv_pct = min(row['Advances'] / row['Advance_Target'] * 100, 100)
            npa_score = max(0, 100 - row['NPA_Percent'] * 15)
            casa_score = min(row['CASA_Percent'] * 2, 100)
            staff_score = min(row['Business_Per_Staff'] / 1.0, 100)

            categories = ['Deposits', 'Advances', 'NPA Health', 'CASA', 'Staff Efficiency']
            values = [dep_pct, adv_pct, npa_score, casa_score, staff_score]
            values_closed = values + [values[0]]
            categories_closed = categories + [categories[0]]

            col1, col2 = st.columns(2)
            with col1:
                # Performance score bars
                fig = go.Figure(go.Bar(
                    x=values, y=categories, orientation='h',
                    marker_color=['#06b6d4','#3b82f6','#10b981','#8b5cf6','#f59e0b'],
                    text=[f"{v:.0f}%" for v in values], textposition='outside',
                    width=0.5
                ))
                fig.update_layout(
                    title=f"📊 {selected} — Performance Breakdown",
                    height=320, plot_bgcolor='white', paper_bgcolor='white',
                    xaxis=dict(range=[0, 120], title="Score (0–100)"),
                    showlegend=False
                )
                st.plotly_chart(fig, key="branch_performance_bar")

            with col2:
                # Deposit vs Advance pie
                fig = go.Figure(data=[go.Pie(
                    labels=['Deposits', 'Advances'],
                    values=[row['Total_Deposits'], row['Advances']],
                    marker_colors=['#06b6d4', '#667eea'],
                    hole=0.4, textinfo='label+percent'
                )])
                fig.update_layout(title=f"💰 {selected} — Business Mix", height=300,
                                  plot_bgcolor='white', paper_bgcolor='white')
                st.plotly_chart(fig, key="branch_business_mix_pie")

            # Insight cards
            st.markdown("---")
            c1, c2, c3 = st.columns(3)
            with c1:
                st.markdown("### 🔍 Areas for Attention")
                if analysis['insights']['needs_support']:
                    for ins in analysis['insights']['needs_support']:
                        st.markdown(f'<div class="alert-info"><strong>{ins["title"]}</strong><br>{ins["detail"]}</div>', unsafe_allow_html=True)
                else:
                    st.markdown('<div class="alert-strength">All good here! ✅</div>', unsafe_allow_html=True)
            with c2:
                st.markdown("### 👀 Worth Monitoring")
                if analysis['insights']['watch_area']:
                    for ins in analysis['insights']['watch_area']:
                        st.markdown(f'<div class="alert-watch"><strong>{ins["title"]}</strong><br>{ins["detail"]}</div>', unsafe_allow_html=True)
                else:
                    st.markdown('<div class="alert-strength">Nothing flagged! 👍</div>', unsafe_allow_html=True)
            with c3:
                st.markdown("### ✨ What's Working")
                if analysis['insights']['strengths']:
                    for ins in analysis['insights']['strengths']:
                        st.markdown(f'<div class="alert-strength"><strong>{ins["title"]}</strong><br>{ins["detail"]}</div>', unsafe_allow_html=True)
                else:
                    st.markdown('<div class="alert-info">Keep working at it — progress takes time! 💪</div>', unsafe_allow_html=True)

        # ════════════════════════════════════════
        # TAB 4 — COMPARE
        # ════════════════════════════════════════
        with tab4:
            st.markdown('<p class="section-header">🔄 Branch Comparison</p>', unsafe_allow_html=True)
            branches = df['Branch_Name'].tolist()
            col1, col2 = st.columns(2)
            with col1: b1 = st.selectbox("Branch A", branches, index=0, key="cmp_a")
            with col2: b2 = st.selectbox("Branch B", branches, index=min(1, len(branches)-1), key="cmp_b")

            if b1 == b2:
                st.warning("Please select two different branches to compare.")
            else:
                r1 = df[df['Branch_Name'] == b1].iloc[0]
                r2 = df[df['Branch_Name'] == b2].iloc[0]

                metrics_cmp = [
                    ("Total Deposits (Cr)", 'Total_Deposits', "₹{:.1f}"),
                    ("Deposit Target (Cr)", 'Deposit_Target', "₹{:.1f}"),
                    ("Deposit Achievement", None, "{:.1f}%"),
                    ("Advances (Cr)", 'Advances', "₹{:.1f}"),
                    ("Advance Achievement", None, "{:.1f}%"),
                    ("NPA %", 'NPA_Percent', "{:.2f}%"),
                    ("CASA %", 'CASA_Percent', "{:.1f}%"),
                    ("CD Ratio", 'CD_Ratio', "{:.1f}%"),
                    ("Staff Count", 'Staff_Count', "{}"),
                    ("Business / Staff", 'Business_Per_Staff', "₹{:.1f} Cr"),
                    ("Profit / Staff", 'Profit_Per_Staff', "₹{:.1f} Cr"),
                ]

                ca, cb = st.columns(2)
                with ca:
                    st.markdown(f'<div class="compare-card"><div class="compare-header" style="color:#06b6d4;">📍 {b1}</div>', unsafe_allow_html=True)
                    for label, key, fmt in metrics_cmp:
                        if key is None:
                            if "Deposit Ach" in label: v1 = r1['Total_Deposits']/r1['Deposit_Target']*100
                            else: v1 = r1['Advances']/r1['Advance_Target']*100
                        else:
                            v1 = r1[key]
                        st.markdown(f'<div class="compare-row"><span class="compare-label">{label}</span><span class="compare-value">{fmt.format(v1)}</span></div>', unsafe_allow_html=True)
                    st.markdown('</div>', unsafe_allow_html=True)

                with cb:
                    st.markdown(f'<div class="compare-card"><div class="compare-header" style="color:#667eea;">📍 {b2}</div>', unsafe_allow_html=True)
                    for label, key, fmt in metrics_cmp:
                        if key is None:
                            if "Deposit Ach" in label: v2 = r2['Total_Deposits']/r2['Deposit_Target']*100
                            else: v2 = r2['Advances']/r2['Advance_Target']*100
                        else:
                            v2 = r2[key]
                        st.markdown(f'<div class="compare-row"><span class="compare-label">{label}</span><span class="compare-value">{fmt.format(v2)}</span></div>', unsafe_allow_html=True)
                    st.markdown('</div>', unsafe_allow_html=True)

                # Side-by-side bar chart
                compare_metrics = ['Total_Deposits', 'Advances', 'NPA_Percent', 'CASA_Percent', 'Business_Per_Staff']
                compare_labels = ['Deposits', 'Advances', 'NPA %', 'CASA %', 'Biz/Staff']
                fig = go.Figure()
                fig.add_trace(go.Bar(name=b1, x=compare_labels, y=[r1[m] for m in compare_metrics], marker_color='#06b6d4'))
                fig.add_trace(go.Bar(name=b2, x=compare_labels, y=[r2[m] for m in compare_metrics], marker_color='#667eea'))
                fig.update_layout(barmode='group', title="📊 Side-by-Side Comparison", height=350,
                                  plot_bgcolor='white', paper_bgcolor='white')
                st.plotly_chart(fig, key="branch_comparison_chart")

        # ════════════════════════════════════════
        # TAB 5 — HEATMAPS
        # ════════════════════════════════════════
        with tab5:
            st.markdown('<p class="section-header">🗺️ Heatmap Visualizations</p>', unsafe_allow_html=True)

            heatmap_metric = st.selectbox("Select Metric:", [
                'NPA_Percent', 'CASA_Percent', 'CD_Ratio', 'Business_Per_Staff', 'Profit_Per_Staff'
            ], format_func=lambda x: {'NPA_Percent':'NPA %','CASA_Percent':'CASA %','CD_Ratio':'CD Ratio',
                                       'Business_Per_Staff':'Business/Staff','Profit_Per_Staff':'Profit/Staff'}.get(x, x))

            hdf = df[['Branch_Name', heatmap_metric]].sort_values(heatmap_metric, ascending=(heatmap_metric != 'NPA_Percent'))

            # Color logic: for NPA lower is better
            if heatmap_metric == 'NPA_Percent':
                colors = ['#10b981' if v <= 3 else '#f59e0b' if v <= 6 else '#ef4444' for v in hdf[heatmap_metric]]
            else:
                # Higher is better for most
                vals = hdf[heatmap_metric].values
                mn, mx = vals.min(), vals.max()
                rng = mx - mn if mx != mn else 1
                colors = []
                for v in vals:
                    pct = (v - mn) / rng
                    if pct >= 0.66: colors.append('#10b981')
                    elif pct >= 0.33: colors.append('#f59e0b')
                    else: colors.append('#ef4444')

            fig = go.Figure(go.Bar(
                x=hdf[heatmap_metric].tolist(),
                y=hdf['Branch_Name'].tolist(),
                orientation='h',
                marker_color=colors,
                text=[f"{v:.2f}" for v in hdf[heatmap_metric]],
                textposition='outside'
            ))
            label_map = {'NPA_Percent':'NPA %','CASA_Percent':'CASA %','CD_Ratio':'CD Ratio',
                         'Business_Per_Staff':'Business/Staff (₹Cr)','Profit_Per_Staff':'Profit/Staff (₹Cr)'}
            fig.update_layout(title=f"🗺️ {label_map.get(heatmap_metric, heatmap_metric)} by Branch",
                              height=380, plot_bgcolor='white', paper_bgcolor='white',
                              xaxis_title=label_map.get(heatmap_metric, heatmap_metric))
            st.plotly_chart(fig, key="branch_heatmap")

            # Legend
            if heatmap_metric == 'NPA_Percent':
                st.markdown("""
                <div class="heatmap-legend">
                    <div class="legend-item"><div class="legend-swatch" style="background:#10b981;"></div> ≤ 3% (Healthy)</div>
                    <div class="legend-item"><div class="legend-swatch" style="background:#f59e0b;"></div> 3–6% (Watch)</div>
                    <div class="legend-item"><div class="legend-swatch" style="background:#ef4444;"></div> > 6% (Attention)</div>
                </div>
                """, unsafe_allow_html=True)
            else:
                st.markdown("""
                <div class="heatmap-legend">
                    <div class="legend-item"><div class="legend-swatch" style="background:#10b981;"></div> Top 33%</div>
                    <div class="legend-item"><div class="legend-swatch" style="background:#f59e0b;"></div> Middle 33%</div>
                    <div class="legend-item"><div class="legend-swatch" style="background:#ef4444;"></div> Bottom 33%</div>
                </div>
                """, unsafe_allow_html=True)

            # Zone-wise grouped heatmap
            st.markdown("### 🗺️ Zone-wise Aggregation")
            zone_agg = df.groupby('Zone').agg({
                'Total_Deposits': 'sum', 'Advances': 'sum',
                'NPA_Percent': 'mean', 'CASA_Percent': 'mean',
                'Staff_Count': 'sum', 'Business_Per_Staff': 'mean'
            }).round(2).reset_index()

            fig = make_subplots(rows=1, cols=2, subplot_titles=("Deposits vs Advances by Zone", "Avg NPA vs CASA by Zone"))
            for i, zone in enumerate(zone_agg['Zone']):
                zr = zone_agg[zone_agg['Zone'] == zone].iloc[0]
                fig.add_trace(go.Bar(name=f"{zone} Dep", x=[zone], y=[zr['Total_Deposits']], marker_color='#06b6d4'), row=1, col=1)
                fig.add_trace(go.Bar(name=f"{zone} Adv", x=[zone], y=[zr['Advances']], marker_color='#667eea'), row=1, col=1)
                fig.add_trace(go.Bar(name=f"{zone} NPA", x=[zone], y=[zr['NPA_Percent']], marker_color='#ef4444'), row=1, col=2)
                fig.add_trace(go.Bar(name=f"{zone} CASA", x=[zone], y=[zr['CASA_Percent']], marker_color='#10b981'), row=1, col=2)

            fig.update_layout(height=320, plot_bgcolor='white', paper_bgcolor='white', barmode='group', showlegend=True)
            st.plotly_chart(fig, key="zone_aggregation_chart")

        # ════════════════════════════════════════
        # TAB 6 — ANOMALIES
        # ════════════════════════════════════════
        with tab6:
            st.markdown('<p class="section-header">🔍 Anomaly Detection</p>', unsafe_allow_html=True)
            anomalies = AnomalyDetector(df).detect()

            if not anomalies:
                st.markdown('<div class="alert-strength" style="text-align:center;padding:2rem;"><h3>👍 All Clear!</h3><p>No unusual patterns detected across any branch or metric.</p></div>', unsafe_allow_html=True)
            else:
                st.markdown(f"Found **{len(anomalies)}** data points that stand out. These aren't necessarily problems — just worth understanding.\n")

                # Summary table
                adf = pd.DataFrame(anomalies)
                adf['Direction'] = adf['direction'].map({'HIGH': '⬆️ HIGH', 'LOW': '⬇️ LOW'})
                st.dataframe(adf[['branch','metric','value','mean','z_score','Direction']].rename(
                    columns={'branch':'Branch','metric':'Metric','value':'Value','mean':'Avg','z_score':'Z-Score'}
                ), use_container_width=True, hide_index=True)

                # Visual
                st.markdown("### 📊 Anomaly Visualization")
                for metric in adf['metric'].unique():
                    sub = adf[adf['metric'] == metric]
                    all_vals = df[metric]
                    fig = go.Figure()
                    fig.add_trace(go.Bar(x=df['Branch_Name'], y=df[metric],
                                         marker_color=['#ef4444' if b in sub['branch'].values else '#06b6d4' for b in df['Branch_Name']],
                                         name=metric))
                    fig.add_hline(y=all_vals.mean(), line_dash="dash", line_color="#374151", annotation_text=f"Avg: {all_vals.mean():.2f}")
                    fig.update_layout(title=f"⚠️ {metric} — Flagged Branches in Red", height=280,
                                      plot_bgcolor='white', paper_bgcolor='white')
                    st.plotly_chart(fig, key=f"anomaly_chart_{metric}")

        # ════════════════════════════════════════
        # TAB 7 — EXPORT
        # ════════════════════════════════════════
        with tab7:
            st.markdown('<p class="section-header">📥 Export & Download</p>', unsafe_allow_html=True)

            st.markdown("""
            <div class="download-section">
                <h3>🎯 Dynamic Excel Dashboard</h3>
                <p style="font-size:1.05rem;margin-bottom:0.8rem;">Generates a fully offline-capable Excel file with:</p>
                <ul style="font-size:0.95rem;line-height:1.8;">
                    <li>✅ Branch dropdown with auto-updating formulas</li>
                    <li>✅ Key metrics, grading, and status indicators</li>
                    <li>✅ Full summary sheet for all branches</li>
                    <li>✅ Professional formatting — share-ready</li>
                </ul>
            </div>
            """, unsafe_allow_html=True)

            if st.button("📊 Generate Excel Dashboard", key="generate_excel_btn", type="primary"):
                with st.spinner("Building your dashboard..."):
                    excel_bytes = create_dynamic_excel(df)
                st.download_button("📥 Download Excel Dashboard", excel_bytes,
                    f"BankVista_Dashboard_{date.today().strftime('%Y%m%d')}.xlsx",
                    "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", key="download_excel_btn")

            st.markdown("---")
            c1, c2 = st.columns(2)
            with c1:
                st.markdown("""
                <div class="feature-card">
                    <div class="feature-icon">💬</div>
                    <div class="feature-title">Export AI Conversation</div>
                    <div class="feature-desc">Download your chat history as JSON</div>
                </div>
                """, unsafe_allow_html=True)
                if st.button("📝 Export Chat", key="export_chat_btn"):
                    st.download_button("📥 Download JSON", st.session_state['ai_assistant'].export_conversation(),
                        f"BankVista_Chat_{date.today().strftime('%Y%m%d')}.json", "application/json", use_container_width=True)
            with c2:
                st.markdown("""
                <div class="feature-card">
                    <div class="feature-icon">📊</div>
                    <div class="feature-title">Export Raw Data</div>
                    <div class="feature-desc">Download current dataset as CSV</div>
                </div>
                """, unsafe_allow_html=True)
                st.download_button("📥 Download CSV", df.to_csv(index=False),
                    f"BankVista_Data_{date.today().strftime('%Y%m%d')}.csv", "text/csv", use_container_width=True)

    else:
        # ── Welcome screen ──
        st.markdown("""
        <div style="text-align:center;padding:3.5rem 2rem;background:white;border-radius:24px;margin-top:1.5rem;box-shadow:0 4px 6px rgba(0,0,0,0.05);">
            <div style="font-size:3.5rem;margin-bottom:0.8rem;">🚀</div>
            <h2 style="font-size:2.2rem;margin-bottom:0.8rem;font-weight:800;color:#1f2937;">Welcome to BankVista AI</h2>
            <p style="font-size:1.2rem;color:#6b7280;margin-bottom:1.5rem;">AI-powered banking analytics that understands plain English</p>
            <p style="font-size:1rem;color:#9ca3af;">Upload your data or click "Try Sample Data" in the sidebar to begin</p>
        </div>
        """, unsafe_allow_html=True)

        st.markdown('<p class="section-header">✨ What Can BankVista AI Do?</p>', unsafe_allow_html=True)
        cards = [
            ("🤖", "Smart AI Chat", "Just type naturally. Say \"which loans are bad?\" and it understands you mean NPA analysis."),
            ("📈", "Predictions", "Forecast NPA trends and target achievement probabilities for any branch."),
            ("🔄", "Branch Compare", "Put any two branches side-by-side with full metrics and visual charts."),
            ("🗺️", "Heatmaps", "Color-coded visualizations of every metric across all branches."),
            ("🔍", "Anomaly Detection", "Automatically flags unusual data points using Z-score analysis."),
            ("📥", "Excel Export", "One-click dynamic Excel dashboards that work completely offline."),
        ]
        cols = st.columns(3)
        for (icon, title, desc), col in zip(cards, cols * 2):
            with col:
                st.markdown(f"""
                <div class="feature-card">
                    <div class="feature-icon">{icon}</div>
                    <div class="feature-title">{title}</div>
                    <div class="feature-desc">{desc}</div>
                </div>
                """, unsafe_allow_html=True)


if __name__ == "__main__":
    main()
