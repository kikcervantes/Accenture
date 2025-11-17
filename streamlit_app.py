import streamlit as st
import os
import re
import pandas as pd
import fitz
from collections import Counter
from fpdf import FPDF
from rapidfuzz import fuzz
import tempfile
import base64
import yfinance as yf
import requests
import plotly.graph_objects as go
import plotly.express as px
from datetime import datetime, timedelta
import json

# ----------------------------- CONFIGURATION -----------------------------
st.set_page_config(
    page_title="Financial Analyzer Pro | Accenture",
    page_icon="📊",
    layout="wide",
    initial_sidebar_state="collapsed"
)

# Custom CSS para estilo Accenture
st.markdown("""
<style>
    .main-header {
        font-size: 2.5rem;
        color: #000000;
        font-weight: 700;
        margin-bottom: 0rem;
    }
    .sub-header {
        font-size: 1.2rem;
        color: #A100FF;
        font-weight: 400;
        margin-bottom: 2rem;
    }
    .upload-section {
        background-color: #f8f9fa;
        padding: 2rem;
        border-radius: 10px;
        border-left: 4px solid #A100FF;
        margin-bottom: 2rem;
    }
    .success-box {
        background-color: #d4edda;
        color: #155724;
        padding: 1rem;
        border-radius: 5px;
        border-left: 4px solid #28a745;
        margin: 1rem 0;
    }
    .info-box {
        background-color: #d1ecf1;
        color: #0c5460;
        padding: 1rem;
        border-radius: 5px;
        border-left: 4px solid #17a2b8;
        margin: 1rem 0;
    }
    .metric-card {
        background-color: white;
        padding: 1rem;
        border-radius: 8px;
        border: 1px solid #e0e0e0;
        box-shadow: 0 2px 4px rgba(0,0,0,0.1);
    }
    .swot-box {
        background-color: white;
        padding: 1.5rem;
        border-radius: 8px;
        border: 1px solid #e0e0e0;
        box-shadow: 0 2px 4px rgba(0,0,0,0.1);
        margin-bottom: 1rem;
    }
    .news-card {
        background-color: white;
        padding: 1rem;
        border-radius: 8px;
        border: 1px solid #e0e0e0;
        box-shadow: 0 2px 4px rgba(0,0,0,0.1);
        margin-bottom: 1rem;
    }
    .news-title {
        font-weight: 600;
        color: #A100FF;
        margin-bottom: 0.5rem;
    }
    .news-source {
        font-size: 0.8rem;
        color: #666;
        margin-bottom: 0.5rem;
    }
    .news-date {
        font-size: 0.8rem;
        color: #999;
    }
    .swot-strengths { border-left: 4px solid #28a745; }
    .swot-weaknesses { border-left: 4px solid #dc3545; }
    .swot-opportunities { border-left: 4px solid #17a2b8; }
    .swot-threats { border-left: 4px solid #ffc107; }
    .stTabs [data-baseweb="tab-list"] {
        gap: 2rem;
    }
    .stTabs [data-baseweb="tab"] {
        height: 50px;
        white-space: pre-wrap;
        background-color: #f8f9fa;
        border-radius: 5px 5px 0px 0px;
        gap: 1rem;
        padding-top: 10px;
        padding-bottom: 10px;
    }
    .download-btn {
        background-color: #A100FF;
        color: white;
        padding: 0.5rem 2rem;
        border: none;
        border-radius: 5px;
        font-weight: 600;
    }
    .download-btn:hover {
        background-color: #8A00E6;
        color: white;
    }
</style>
""", unsafe_allow_html=True)

# ----------------------------- IMPROVED NEWS API INTEGRATION -----------------------------

def get_risk_news(company_name, api_key=None):
    """Get relevant risk news for the company using multiple approaches"""
    try:
        if not api_key:
            api_key = "e35e621ac9a245a8879ef9a2505f2646"
        
        # Limpiar el nombre de la compañía para búsqueda
        clean_company_name = re.sub(r'[^\w\s]', '', company_name).strip()
        
        # Si el nombre es muy largo, tomar las primeras palabras
        name_parts = clean_company_name.split()
        if len(name_parts) > 3:
            search_name = " ".join(name_parts[:3])
        else:
            search_name = clean_company_name
        
        search_queries = [
            f"{search_name} financial risk",
            f"{search_name} regulatory compliance", 
            f"{search_name} banking risk",
            f"{search_name} financial compliance",
            f"{search_name} AML KYC",
            search_name,  # Búsqueda general
            f'"{search_name}" company',  # Búsqueda exacta
        ]
        
        all_articles = []
        
        for query in search_queries:
            try:
                # Calculate date range (last 30 days)
                to_date = datetime.now()
                from_date = to_date - timedelta(days=30)
                
                url = "https://newsapi.org/v2/everything"
                
                params = {
                    'q': query,
                    'from': from_date.strftime('%Y-%m-%d'),
                    'to': to_date.strftime('%Y-%m-%d'),
                    'sortBy': 'relevancy',  # Cambiado a relevancy para mejores resultados
                    'language': 'en',
                    'pageSize': 10,
                    'apiKey': api_key
                }
                
                response = requests.get(url, params=params, timeout=15)
                
                if response.status_code == 200:
                    data = response.json()
                    if data['status'] == 'ok' and data['totalResults'] > 0:
                        for article in data['articles']:
                            # Verificar duplicados por título
                            title = article.get('title', '')
                            if not any(a.get('title') == title for a in all_articles):
                                all_articles.append(article)
                        
                        if len(all_articles) >= 15:
                            break
                elif response.status_code == 426:  # Upgrade Required
                    st.sidebar.warning("News API requires upgraded plan for multiple queries")
                    break
                    
            except Exception as e:
                continue
        
        return all_articles[:15]
        
    except Exception as e:
        st.sidebar.warning(f"Could not fetch news: {str(e)}")
        return []

def display_news_articles(articles, company_name):
    """Display news articles in a formatted way"""
    if not articles:
        st.warning(f"No recent news found for {company_name}. This could be due to:")
        st.info("""
        - API limit reached
        - No recent news about this company
        - Company name might not match news sources
        - Try using the company's stock ticker instead
        """)
        
        # Opción para buscar manualmente
        col1, col2 = st.columns([3, 1])
        with col1:
            manual_search = st.text_input("Try alternative search term:", 
                                        value=company_name,
                                        key="manual_news_search")
        with col2:
            if st.button("🔍 Search News", key="manual_news_btn"):
                with st.spinner("Searching..."):
                    news_api_key = "e35e621ac9a245a8879ef9a2505f2646"
                    new_articles = get_risk_news(manual_search, news_api_key)
                    if new_articles:
                        st.success(f"Found {len(new_articles)} articles!")
                        st.rerun()
                    else:
                        st.error("No articles found with this search term.")
        return
    
    st.subheader(f"📰 Recent News for {company_name}")
    st.info(f"Found {len(articles)} recent articles")
    
    for i, article in enumerate(articles):
        with st.container():
            col1, col2 = st.columns([4, 1])
            
            with col1:
                # Título con enlace
                title = article.get("title", "No title")
                url = article.get("url", "")
                
                if url:
                    st.markdown(f'<div class="news-title"><a href="{url}" target="_blank">{title}</a></div>', unsafe_allow_html=True)
                else:
                    st.markdown(f'<div class="news-title">{title}</div>', unsafe_allow_html=True)
                
                # Fuente
                source = article.get("source", {}).get("name", "Unknown source")
                st.markdown(f'<div class="news-source">Source: {source}</div>', unsafe_allow_html=True)
                
                # Descripción
                description = article.get("description", "")
                if description:
                    st.write(description[:200] + "..." if len(description) > 200 else description)
            
            with col2:
                # Fecha
                if article.get('publishedAt'):
                    try:
                        published_date = datetime.strptime(article['publishedAt'][:10], '%Y-%m-%d')
                        st.markdown(f'<div class="news-date">{published_date.strftime("%b %d, %Y")}</div>', unsafe_allow_html=True)
                    except:
                        st.markdown(f'<div class="news-date">{article["publishedAt"][:10]}</div>', unsafe_allow_html=True)
                
                # Imagen (si existe)
                if article.get('urlToImage'):
                    st.image(article['urlToImage'], width=80)
            
            if i < len(articles) - 1:
                st.markdown("---")

# ----------------------------- EXTRACT COMPANY NAME FROM ANNUAL REPORT -----------------------------

def extract_company_name_from_annual(text):
    """Extract company name from annual report text"""
    try:
        # Patrones comunes para encontrar el nombre de la compañía en reports anuales
        patterns = [
            r"annual report of\s+([^\n,.]+)",
            r"consolidated financial statements of\s+([^\n,.]+)",
            r"company name[:\s]+([^\n,.]+)",
            r"report of\s+([^\n,.]+)",
            r"financial statements\s+([^\n,.]+)",
            r"([A-Z][A-Za-z0-9&\.\-\s]+(?:Inc|Ltd|Corporation|Company|Group|PLC|NV|SA)[\.]?)",
        ]
        
        for pattern in patterns:
            matches = re.findall(pattern, text, re.IGNORECASE)
            if matches:
                for match in matches:
                    if isinstance(match, tuple):
                        name = match[0].strip()
                    else:
                        name = match.strip()
                    
                    if len(name) > 3 and len(name) < 100:  # Validar que sea un nombre razonable
                        # Filtrar falsos positivos
                        if not any(word in name.lower() for word in ['balance', 'income', 'cash', 'statement', 'report', 'consolidated']):
                            return name
        
        # Buscar líneas que parezcan nombres de compañías (todo en mayúsculas)
        lines = text.split('\n')
        for line in lines:
            line = line.strip()
            # Los nombres de compañías en reports suelen estar en mayúsculas o tener formato específico
            if (len(line) > 5 and len(line) < 80 and 
                not any(word in line.lower() for word in ['balance', 'income', 'cash', 'statement', 'report', 'notes', 'page']) and
                (line.isupper() or any(keyword in line.lower() for keyword in ['inc', 'ltd', 'corp', 'company', 'group']))):
                return line
        
        return None
        
    except Exception as e:
        return None

# ----------------------------- IMPROVED KVK EXTRACTION -----------------------------

FIELDS_TO_EXTRACT = [
    "Name",
    "Chamber of Commerce number",
    "RSIN",
    "Legal form",
    "Statutory seat"
]

KNOWN_LABELS = set(l.strip().lower() for l in FIELDS_TO_EXTRACT + [
    "Additional information", "Branch", "Establishment number", "Trade name",
    "Visiting address", "Previous (statutory) names", "Previous trade names",
    "Previous addresses", "Previous activity descriptions", "Other history",
    "Transfers of the establishment", "Takeover of", "rechtspersoon", "onderneming",
    "vestiging", "legal entity", "company", "establishment"
])

def load_pdf_lines(pdf_path):
    """Load PDF and return lines"""
    doc = fitz.open(pdf_path)
    lines = []
    for page in doc:
        lines.extend([line.strip() for line in page.get_text().splitlines() if line.strip()])
    return lines

def extract_kvk_method_block(lines):
    """Original block matching - most reliable for standard KVK format"""
    def find_field_block(lines, fields, use_fuzzy=True):
        target_fields = [f.strip().lower() for f in fields]
        for i in range(len(lines) - len(fields)):
            candidate = [lines[j].strip().lower() for j in range(i, i + len(fields))]
            if use_fuzzy:
                match_scores = [fuzz.ratio(a, b) for a, b in zip(candidate, target_fields)]
                if all(score > 85 for score in match_scores):
                    return i
            else:
                if candidate == target_fields:
                    return i
        return None

    def extract_value_block(lines, start_index, fields):
        for i in range(start_index + len(fields), len(lines) - len(fields)):
            candidate = lines[i:i + len(fields)]
            if all(line.strip().lower() not in KNOWN_LABELS for line in candidate):
                return candidate
        return []

    field_block_index = find_field_block(lines, FIELDS_TO_EXTRACT, use_fuzzy=True)
    if field_block_index is not None:
        value_block = extract_value_block(lines, field_block_index, FIELDS_TO_EXTRACT)
        if len(value_block) == len(FIELDS_TO_EXTRACT):
            return dict(zip(FIELDS_TO_EXTRACT, value_block))
    return {}

def extract_kvk_method_keyvalue(lines):
    """Extract using key-value pattern matching"""
    result = {}
    
    # Map field names to their possible labels
    field_keywords = {
        "Name": ["statutaire naam", "statutory name", "name given in the articles", "handelsnaam"],
        "Chamber of Commerce number": ["kvk-nummer", "cci number", "chamber of commerce"],
        "RSIN": ["rsin"],
        "Legal form": ["rechtsvorm", "legal form"],
        "Statutory seat": ["statutaire zetel", "corporate seat", "statutory seat"]
    }
    
    i = 0
    while i < len(lines):
        line = lines[i].strip()
        line_lower = line.lower()
        
        # Check each field
        for field_name, keywords in field_keywords.items():
            if field_name in result:
                continue
                
            # Check if any keyword is in this line
            for keyword in keywords:
                if keyword in line_lower:
                    # Try to extract from same line first
                    if ":" in line or "  " in line:
                        parts = re.split(r'[:\t]+', line, 1)
                        if len(parts) > 1:
                            value = parts[1].strip()
                            if value and value.lower() not in KNOWN_LABELS:
                                result[field_name] = clean_kvk_value(value, field_name)
                                break
                    
                    # Check next lines
                    if field_name not in result:
                        for j in range(i + 1, min(i + 5, len(lines))):
                            next_line = lines[j].strip()
                            if next_line and next_line.lower() not in KNOWN_LABELS:
                                cleaned = clean_kvk_value(next_line, field_name)
                                if cleaned != "Not found":
                                    result[field_name] = cleaned
                                    break
                    break
        i += 1
    
    return result

def clean_kvk_value(value, field_type):
    """Clean extracted KVK values"""
    if not value:
        return "Not found"
    
    value = value.strip()
    
    # Special handling for specific fields
    if field_type == "Chamber of Commerce number":
        # Extract 8-digit number
        match = re.search(r'\b(\d{8})\b', value)
        return match.group(1) if match else "Not found"
    
    elif field_type == "RSIN":
        # Extract 9-digit number
        match = re.search(r'\b(\d{9})\b', value)
        return match.group(1) if match else "Not found"
    
    elif field_type == "Name":
        # Remove common prefixes
        value = re.sub(r'^(Statutory name|Statutaire naam|Name given in the articles)[:\s]+', '', value, flags=re.I)
        if len(value) < 2 or len(value) > 150:
            return "Not found"
        return value
    
    elif field_type == "Legal form":
        # Standardize legal forms
        if re.search(r'\bB\.?V\.?\b', value, re.I) or "besloten vennootschap" in value.lower():
            return "Besloten Vennootschap (B.V.)"
        elif re.search(r'\bN\.?V\.?\b', value, re.I) or "naamloze vennootschap" in value.lower():
            return "Naamloze Vennootschap (N.V.)"
        elif "private limited" in value.lower():
            return "Private Limited Company"
        return value
    
    elif field_type == "Statutory seat":
        # Extract just the city name
        match = re.search(r'\b([A-Z][a-z]+(?:\s+[A-Z][a-z]+)?)\b', value)
        return match.group(1) if match else value
    
    return value if value else "Not found"

def extract_kvk_data_robust(pdf_path):
    """Extract KVK data with fallback methods"""
    lines = load_pdf_lines(pdf_path)
    
    # Show debug in sidebar
    with st.sidebar.expander("📋 KVK Extraction Debug"):
        st.write("First 50 lines:")
        st.text("\n".join(lines[:50]))
    
    # Try block method first (most reliable)
    result1 = extract_kvk_method_block(lines)
    score1 = sum(1 for v in result1.values() if v and v != "Not found")
    
    # Try key-value method as fallback
    result2 = extract_kvk_method_keyvalue(lines)
    score2 = sum(1 for v in result2.values() if v and v != "Not found")
    
    # Merge results - prefer block method, fill gaps with key-value
    final_result = {}
    for field in FIELDS_TO_EXTRACT:
        if field in result1 and result1[field] != "Not found":
            final_result[field] = result1[field]
        elif field in result2 and result2[field] != "Not found":
            final_result[field] = result2[field]
        else:
            final_result[field] = "Not found"
    
    final_score = sum(1 for v in final_result.values() if v and v != "Not found")
    
    return pd.DataFrame(list(final_result.items()), columns=["Field", "Value"])

# ----------------------------- ANNUAL REPORT EXTRACTION -----------------------------

def extract_years(text):
    """Extract most common years"""
    all_years = re.findall(r"\b20\d{2}\b", text)
    most_common = Counter(all_years).most_common(2)
    return [y for y, _ in most_common]

def parse_numeric(s):
    """Parse numeric value from string"""
    if not s:
        return None
    # Remover caracteres especiales que causan problemas en PDF
    s = str(s).replace(",", "").replace("€", "EUR ").replace("$", "USD ").replace(" ", "").strip()
    if "(" in s and ")" in s:
        s = "-" + s.strip("()")
    try:
        return float(s)
    except:
        return None

def safe_divide(numerator, denominator, ratio_name, year):
    """Safe division for ratios"""
    if pd.isna(numerator) or pd.isna(denominator) or denominator == 0:
        return None
    return numerator / denominator

# ----------------------------- BALANCE SHEET -----------------------------

def extract_balance_sheet(text, years):
    """Extract balance sheet - IMPROVED with better error handling"""
    bc_labels = {
        "Current Assets": ["Total current assets", "Current assets"],
        "Fixed Assets": ["Total non-current assets", "Fixed assets", "Non-current assets"],
        "Total Assets": ["Total assets"],
        "Current Liabilities": ["Total current liabilities", "Current liabilities"],
        "Long-Term Liabilities": ["Total non-current liabilities", "Long-term liabilities", "Non-current liabilities"],
        "Equity": ["Total equity", "Equity", "Shareholders' equity"]
    }

    def extract_value(labels, text, years):
        """Try multiple label variations"""
        for label in labels:
            # Try exact pattern with two numbers
            pattern = rf"{re.escape(label)}\s+([\d,]+)\s+([\d,]+)"
            matches = re.findall(pattern, text, re.IGNORECASE)
            if matches:
                last_match = matches[-1]
                return {
                    years[0]: last_match[0].replace(',', ''),
                    years[1]: last_match[1].replace(',', '')
                }
        return {years[0]: None, years[1]: None}

    data = {}
    for key, label_list in bc_labels.items():
        data[key] = extract_value(label_list, text, years)

    df = pd.DataFrame(data).T
    if df.empty:
        return pd.DataFrame(columns=["Items", years[0], years[1]])
    
    df.columns = years
    df = df.reset_index().rename(columns={"index": "Items"})
    return df

# ----------------------------- PROFIT & LOSS -----------------------------

def find_profit_loss_page(doc):
    """Find P&L page"""
    keywords = ["consolidated income statement", "revenue", "operating profit", "profit for the year"]
    for i in range(len(doc)):
        text = doc.load_page(i).get_text("text").lower()
        if all(kw in text for kw in keywords[:2]) and any(kw in text for kw in keywords[2:]):
            return i
    return 0

def extract_profit_loss(doc, page_number, years):
    """Extract P&L - IMPROVED with better error handling"""
    try:
        page = doc.load_page(page_number)
        lines = [line.strip() for line in page.get_text("text").splitlines()]
    except:
        return pd.DataFrame(columns=["Items", years[0], years[1]])
    
    item_synonyms = {
        "Revenue": ["revenue", "sales", "turnover"],
        "Excise tax expense": ["excise tax expense", "excise duty"],
        "Total other expenses": ["total other expenses", "other expenses"],
        "Operating profit": ["operating profit", "operating income"],
        "Net finance expenses": ["net finance expenses", "finance costs", "interest expense"],
        "Share of profit of associates": ["share of profit of associates", "joint ventures"],
        "Profit before income tax": ["profit before income tax", "income before taxes"],
        "Income tax expense": ["income tax expense", "tax expense"],
        "Profit for the year": ["profit for the year", "net profit", "net income", "profit"]
    }

    def find_two_numbers_after(index, lines):
        found_vals = []
        for offset in range(1, 6):
            if index + offset < len(lines):
                nums = re.findall(r"[-(]?\d[\d,.)]+", lines[index + offset])
                for n in nums:
                    num = parse_numeric(n)
                    if num is not None and abs(num) >= 100:
                        found_vals.append(num)
                        if len(found_vals) == 2:
                            return found_vals[0], found_vals[1]
        return None, None

    data_rows = []
    used_items = set()
    
    for i, line in enumerate(lines):
        lcline = line.lower()
        for item_name, synonyms in item_synonyms.items():
            if item_name in used_items:
                continue
            if any(syn in lcline for syn in synonyms):
                val1, val2 = find_two_numbers_after(i, lines)
                if val1 is not None and val2 is not None:
                    data_rows.append({
                        "Items": item_name,
                        years[0]: val1,
                        years[1]: val2
                    })
                    used_items.add(item_name)
                    break
    
    if not data_rows:
        return pd.DataFrame(columns=["Items", years[0], years[1]])
    
    return pd.DataFrame(data_rows)

# ----------------------------- CASH FLOW -----------------------------

def find_cash_flow_page(doc):
    """Find cash flow page"""
    keywords = [
        "cash flow from operating activities",
        "cash flow used in investing",
        "cash flows generated from financing",
        "cash and cash equivalents as at 31 december"
    ]
    for i in range(len(doc)):
        text = doc.load_page(i).get_text("text").lower()
        if any(kw in text for kw in keywords):
            return i
    return min(1, len(doc) - 1)

def extract_cash_flow(doc, page_number, years):
    """Extract cash flow - IMPROVED with better error handling"""
    try:
        page = doc.load_page(page_number)
        lines = [line.strip() for line in page.get_text("text").splitlines() if line.strip()]
    except:
        return pd.DataFrame(columns=["Items", years[0], years[1]])

    field_names = {
        "Cash flow from operating activities": ["cash flow from operating activities"],
        "Cash flow used in investing activities": ["cash flow used in investing activities", "cash flow from investing activities"],
        "Cash flows from financing activities": ["cash flows generated from financing activities", "cash flow from financing activities"],
        "Cash and cash equivalents": ["cash and cash equivalents as at 31 december", "cash and cash equivalents"]
    }

    def find_two_numbers_after(index, lines):
        found_vals = []
        for offset in range(1, 6):
            if index + offset < len(lines):
                nums = re.findall(r"[-(]?\d[\d,.)]+", lines[index + offset])
                for n in nums:
                    num = parse_numeric(n)
                    if num is not None and abs(num) >= 100:
                        found_vals.append(num)
                        if len(found_vals) == 2:
                            return found_vals[0], found_vals[1]
        return None, None

    data_rows = []
    used_fields = set()

    for i, line in enumerate(lines):
        lcline = line.lower()
        for label, synonyms in field_names.items():
            if label in used_fields:
                continue
            if any(syn in lcline for syn in synonyms):
                val1, val2 = find_two_numbers_after(i, lines)
                if val1 is not None or val2 is not None:
                    data_rows.append({
                        "Items": label,
                        years[0]: val1 if val1 is not None else "",
                        years[1]: val2 if val2 is not None else ""
                    })
                    used_fields.add(label)
                    break

    if not data_rows:
        return pd.DataFrame(columns=["Items", years[0], years[1]])
    
    return pd.DataFrame(data_rows)

# ----------------------------- YAHOO FINANCE EXTRACTION -----------------------------

def find_company_ticker(company_name):
    """Find Yahoo Finance ticker symbol for a company"""
    try:
        # Search for ticker using Yahoo Finance
        search_url = f"https://query2.finance.yahoo.com/v1/finance/search?q={company_name}"
        headers = {
            'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/91.0.4472.124 Safari/537.36'
        }
        
        response = requests.get(search_url, headers=headers)
        if response.status_code == 200:
            data = response.json()
            if 'quotes' in data and len(data['quotes']) > 0:
                # Prefer equity quotes
                for quote in data['quotes']:
                    if quote.get('quoteType') in ['EQUITY', 'ETF']:
                        return quote.get('symbol')
                # Return first symbol if no equity found
                return data['quotes'][0].get('symbol')
    except Exception as e:
        st.sidebar.warning(f"Could not find ticker automatically: {str(e)}")
    
    return None

def get_yahoo_finance_data(ticker, company_name):
    """Extract financial data from Yahoo Finance"""
    try:
        stock = yf.Ticker(ticker)
        info = stock.info
        
        # Get financial statements
        balance_sheet = stock.balance_sheet
        income_stmt = stock.income_stmt
        cash_flow = stock.cashflow
        
        if balance_sheet.empty or income_stmt.empty:
            st.error(f"❌ No financial data found for {ticker}")
            return None, None, None, None, None
        
        # Get the two most recent years
        years = balance_sheet.columns[:2].strftime('%Y').tolist()
        
        # Extract Balance Sheet data
        balance_data = {
            "Current Assets": extract_financial_value(balance_sheet, ["Current Assets", "Total Current Assets"]),
            "Fixed Assets": extract_financial_value(balance_sheet, ["Non Current Assets", "Property Plant Equipment Net", "Net PPE"]),
            "Total Assets": extract_financial_value(balance_sheet, ["Total Assets"]),
            "Current Liabilities": extract_financial_value(balance_sheet, ["Current Liabilities", "Total Current Liabilities"]),
            "Long-Term Liabilities": extract_financial_value(balance_sheet, ["Non Current Liabilities", "Long Term Debt"]),
            "Equity": extract_financial_value(balance_sheet, ["Stockholders Equity", "Total Equity"])
        }
        
        # Extract Profit & Loss data
        profit_loss_data = {
            "Revenue": extract_financial_value(income_stmt, ["Total Revenue", "Revenue"]),
            "Operating profit": extract_financial_value(income_stmt, ["Operating Income", "Operating Profit"]),
            "Net finance expenses": extract_financial_value(income_stmt, ["Interest Expense", "Net Interest Income"]),
            "Profit before income tax": extract_financial_value(income_stmt, ["Pretax Income", "Income Before Tax"]),
            "Income tax expense": extract_financial_value(income_stmt, ["Tax Provision", "Income Tax Expense"]),
            "Profit for the year": extract_financial_value(income_stmt, ["Net Income", "Net Income From Continuing Operations"])
        }
        
        # Extract Cash Flow data
        cash_flow_data = {
            "Cash flow from operating activities": extract_financial_value(cash_flow, ["Operating Cash Flow", "Cash From Operating Activities"]),
            "Cash flow used in investing activities": extract_financial_value(cash_flow, ["Investing Cash Flow", "Cash From Investing Activities"]),
            "Cash flows from financing activities": extract_financial_value(cash_flow, ["Financing Cash Flow", "Cash From Financing Activities"]),
            "Cash and cash equivalents": extract_financial_value(balance_sheet, ["Cash And Cash Equivalents", "Cash"])
        }
        
        # Create DataFrames
        bc_df = create_financial_df(balance_data, years, "Balance Sheet")
        pl_df = create_financial_df(profit_loss_data, years, "Profit & Loss")
        cash_df = create_financial_df(cash_flow_data, years, "Cash Flow")
        
        # Get company info
        company_info = {
            "Company Name": info.get('longName', company_name),
            "Sector": info.get('sector', 'N/A'),
            "Industry": info.get('industry', 'N/A'),
            "Market Cap": f"${info.get('marketCap', 0):,}" if info.get('marketCap') else 'N/A',
            "Country": info.get('country', 'N/A'),
            "Currency": info.get('currency', 'USD')
        }
        
        return bc_df, pl_df, cash_df, pd.DataFrame(list(company_info.items()), columns=["Field", "Value"]), years
        
    except Exception as e:
        st.error(f"❌ Error fetching data from Yahoo Finance: {str(e)}")
        return None, None, None, None, None

def extract_financial_value(financial_stmt, possible_keys):
    """Extract financial values using multiple possible keys"""
    for key in possible_keys:
        if key in financial_stmt.index:
            values = financial_stmt.loc[key]
            if len(values) >= 2:
                return {
                    'year1': values.iloc[0] if not pd.isna(values.iloc[0]) else 0,
                    'year2': values.iloc[1] if len(values) > 1 and not pd.isna(values.iloc[1]) else 0
                }
    return {'year1': 0, 'year2': 0}

def create_financial_df(data_dict, years, statement_type):
    """Create financial statement DataFrame"""
    rows = []
    for item, values in data_dict.items():
        rows.append({
            "Items": item,
            years[0]: values['year1'],
            years[1]: values['year2']
        })
    return pd.DataFrame(rows)

# ----------------------------- CORRECTED RATIOS CALCULATION -----------------------------

def compute_ratios(bc_df, pl_df, years):
    """Calculate financial ratios with corrected formulas"""
    if bc_df.empty or pl_df.empty:
        return pd.DataFrame()
    
    # Convert to numeric - CORREGIDO: asegurar que todos los valores sean numéricos
    bc_df_cal = bc_df.set_index('Items').apply(pd.to_numeric, errors='coerce')
    pl_df_cal = pl_df.set_index('Items').apply(pd.to_numeric, errors='coerce')

    ratios = {
        'Current Ratio': [],
        'Quick Ratio': [],
        'Debt to Equity Ratio': [],
        'Return on Equity (ROE)': [],
        'Return on Assets (ROA)': [],
        'Gross Profit Margin': [],
        'Net Profit Margin': [],
        'Asset Turnover': []
    }

    for year in years:
        try:
            # Extract values from balance sheet
            current_assets = bc_df_cal.loc['Current Assets', year] if 'Current Assets' in bc_df_cal.index else None
            current_liabilities = bc_df_cal.loc['Current Liabilities', year] if 'Current Liabilities' in bc_df_cal.index else None
            total_assets = bc_df_cal.loc['Total Assets', year] if 'Total Assets' in bc_df_cal.index else None
            total_equity = bc_df_cal.loc['Equity', year] if 'Equity' in bc_df_cal.index else None
            long_term_liabilities = bc_df_cal.loc['Long-Term Liabilities', year] if 'Long-Term Liabilities' in bc_df_cal.index else None
            
            # Extract values from profit & loss
            revenue = pl_df_cal.loc['Revenue', year] if 'Revenue' in pl_df_cal.index else None
            operating_profit = pl_df_cal.loc['Operating profit', year] if 'Operating profit' in pl_df_cal.index else None
            net_income = pl_df_cal.loc['Profit for the year', year] if 'Profit for the year' in pl_df_cal.index else None
            
            # Calculate total liabilities
            total_liabilities = None
            if current_liabilities is not None and long_term_liabilities is not None:
                total_liabilities = current_liabilities + long_term_liabilities
            
            # Calculate ratios with corrected formulas
            current_ratio = safe_divide(current_assets, current_liabilities, 'Current Ratio', year)
            
            # Quick Ratio (assuming inventory is part of current assets - simplified)
            quick_ratio = safe_divide(current_assets, current_liabilities, 'Quick Ratio', year)
            
            debt_to_equity = safe_divide(total_liabilities, total_equity, 'Debt to Equity Ratio', year)
            roe = safe_divide(net_income, total_equity, 'Return on Equity', year)
            roa = safe_divide(net_income, total_assets, 'Return on Assets', year)
            
            # Profit margins
            gross_profit_margin = safe_divide(operating_profit, revenue, 'Gross Profit Margin', year)
            net_profit_margin = safe_divide(net_income, revenue, 'Net Profit Margin', year)
            
            # Asset turnover
            asset_turnover = safe_divide(revenue, total_assets, 'Asset Turnover', year)
            
            ratios['Current Ratio'].append(current_ratio)
            ratios['Quick Ratio'].append(quick_ratio)
            ratios['Debt to Equity Ratio'].append(debt_to_equity)
            ratios['Return on Equity (ROE)'].append(roe)
            ratios['Return on Assets (ROA)'].append(roa)
            ratios['Gross Profit Margin'].append(gross_profit_margin)
            ratios['Net Profit Margin'].append(net_profit_margin)
            ratios['Asset Turnover'].append(asset_turnover)
            
        except Exception as e:
            # If any error occurs, append None for all ratios for this year
            for ratio in ratios:
                ratios[ratio].append(None)

    df_ratios = pd.DataFrame(ratios, index=years).T.reset_index()
    df_ratios.rename(columns={"index": "Ratios"}, inplace=True)
    return df_ratios

# ----------------------------- FINANCIAL CHARTS -----------------------------

def create_financial_charts(bc_df, pl_df, cash_df, ratios_df, years):
    """Create comprehensive financial charts"""
    charts = {}
    
    # 1. Revenue vs Profit Trend
    if not pl_df.empty:
        try:
            revenue_data = pl_df[pl_df['Items'] == 'Revenue']
            profit_data = pl_df[pl_df['Items'] == 'Profit for the year']
            
            if not revenue_data.empty and not profit_data.empty:
                fig_revenue_profit = go.Figure()
                
                # Revenue bars
                fig_revenue_profit.add_trace(go.Bar(
                    name='Revenue',
                    x=years,
                    y=[revenue_data[year].iloc[0] for year in years],
                    marker_color='#A100FF'
                ))
                
                # Profit line
                fig_revenue_profit.add_trace(go.Scatter(
                    name='Net Profit',
                    x=years,
                    y=[profit_data[year].iloc[0] for year in years],
                    mode='lines+markers',
                    line=dict(color='#00FF88', width=3),
                    yaxis='y2'
                ))
                
                fig_revenue_profit.update_layout(
                    title='Revenue vs Net Profit Trend',
                    xaxis_title='Year',
                    yaxis_title='Revenue',
                    yaxis2=dict(title='Net Profit', overlaying='y', side='right'),
                    barmode='group',
                    height=400
                )
                charts['revenue_profit'] = fig_revenue_profit
        except:
            pass
    
    # 2. Balance Sheet Composition
    if not bc_df.empty:
        try:
            balance_items = ['Current Assets', 'Fixed Assets', 'Current Liabilities', 'Long-Term Liabilities', 'Equity']
            available_items = [item for item in balance_items if item in bc_df['Items'].values]
            
            if available_items:
                fig_balance = go.Figure()
                
                for item in available_items:
                    item_data = bc_df[bc_df['Items'] == item]
                    fig_balance.add_trace(go.Bar(
                        name=item,
                        x=years,
                        y=[item_data[year].iloc[0] for year in years]
                    ))
                
                fig_balance.update_layout(
                    title='Balance Sheet Composition',
                    xaxis_title='Year',
                    yaxis_title='Amount',
                    barmode='stack',
                    height=400
                )
                charts['balance_sheet'] = fig_balance
        except:
            pass
    
    # 3. Profitability Ratios
    if not ratios_df.empty:
        try:
            profit_ratios = ['Gross Profit Margin', 'Net Profit Margin', 'Return on Assets (ROA)', 'Return on Equity (ROE)']
            available_ratios = [ratio for ratio in profit_ratios if ratio in ratios_df['Ratios'].values]
            
            if available_ratios:
                fig_profitability = go.Figure()
                
                for ratio in available_ratios:
                    ratio_data = ratios_df[ratios_df['Ratios'] == ratio]
                    fig_profitability.add_trace(go.Scatter(
                        name=ratio,
                        x=years,
                        y=[ratio_data[year].iloc[0] for year in years],
                        mode='lines+markers'
                    ))
                
                fig_profitability.update_layout(
                    title='Profitability Ratios Trend',
                    xaxis_title='Year',
                    yaxis_title='Ratio Value',
                    height=400
                )
                charts['profitability'] = fig_profitability
        except:
            pass
    
    # 4. Liquidity and Solvency Ratios
    if not ratios_df.empty:
        try:
            liquidity_ratios = ['Current Ratio', 'Quick Ratio', 'Debt to Equity Ratio']
            available_liquidity = [ratio for ratio in liquidity_ratios if ratio in ratios_df['Ratios'].values]
            
            if available_liquidity:
                fig_liquidity = go.Figure()
                
                for ratio in available_liquidity:
                    ratio_data = ratios_df[ratios_df['Ratios'] == ratio]
                    fig_liquidity.add_trace(go.Scatter(
                        name=ratio,
                        x=years,
                        y=[ratio_data[year].iloc[0] for year in years],
                        mode='lines+markers'
                    ))
                
                fig_liquidity.update_layout(
                    title='Liquidity & Solvency Ratios',
                    xaxis_title='Year',
                    yaxis_title='Ratio Value',
                    height=400
                )
                charts['liquidity'] = fig_liquidity
        except:
            pass
    
    return charts

# ----------------------------- SWOT ANALYSIS -----------------------------

def generate_swot_analysis(bc_df, pl_df, ratios_df, years, company_name):
    """Generate SWOT analysis based on financial data"""
    swot = {
        'Strengths': [],
        'Weaknesses': [],
        'Opportunities': [],
        'Threats': []
    }
    
    try:
        # Get latest year data
        latest_year = years[0]
        
        # Extract key metrics
        current_ratio = None
        debt_to_equity = None
        roe = None
        net_margin = None
        revenue_growth = None
        
        if not ratios_df.empty:
            current_ratio_data = ratios_df[ratios_df['Ratios'] == 'Current Ratio']
            if not current_ratio_data.empty:
                current_ratio = current_ratio_data[latest_year].iloc[0]
            
            debt_data = ratios_df[ratios_df['Ratios'] == 'Debt to Equity Ratio']
            if not debt_data.empty:
                debt_to_equity = debt_data[latest_year].iloc[0]
            
            roe_data = ratios_df[ratios_df['Ratios'] == 'Return on Equity (ROE)']
            if not roe_data.empty:
                roe = roe_data[latest_year].iloc[0]
            
            margin_data = ratios_df[ratios_df['Ratios'] == 'Net Profit Margin']
            if not margin_data.empty:
                net_margin = margin_data[latest_year].iloc[0]
        
        # Calculate revenue growth if we have two years
        if len(years) >= 2 and not pl_df.empty:
            revenue_data = pl_df[pl_df['Items'] == 'Revenue']
            if not revenue_data.empty:
                rev_current = revenue_data[years[0]].iloc[0]
                rev_previous = revenue_data[years[1]].iloc[0]
                if rev_previous and rev_previous != 0:
                    revenue_growth = (rev_current - rev_previous) / rev_previous
        
        # Strengths
        if current_ratio and current_ratio > 2.0:
            swot['Strengths'].append("Strong liquidity position with high current ratio")
        if roe and roe > 0.15:
            swot['Strengths'].append("High return on equity indicating efficient use of shareholder capital")
        if net_margin and net_margin > 0.1:
            swot['Strengths'].append("Healthy profit margins compared to industry standards")
        if revenue_growth and revenue_growth > 0.1:
            swot['Strengths'].append("Strong revenue growth trajectory")
        
        # Weaknesses
        if current_ratio and current_ratio < 1.0:
            swot['Weaknesses'].append("Potential liquidity concerns with current ratio below 1")
        if debt_to_equity and debt_to_equity > 2.0:
            swot['Weaknesses'].append("High debt levels relative to equity")
        if roe and roe < 0.05:
            swot['Weaknesses'].append("Low return on equity indicating inefficient capital utilization")
        if net_margin and net_margin < 0.05:
            swot['Weaknesses'].append("Thin profit margins affecting overall profitability")
        
        # Opportunities (more general)
        swot['Opportunities'].append("Potential for market expansion in current sector")
        swot['Opportunities'].append("Digital transformation initiatives to improve efficiency")
        swot['Opportunities'].append("Strategic partnerships or acquisitions")
        if revenue_growth and revenue_growth > 0:
            swot['Opportunities'].append("Positive growth momentum to leverage for expansion")
        
        # Threats (more general)
        swot['Threats'].append("Economic uncertainty and market volatility")
        swot['Threats'].append("Increasing competition in the industry")
        swot['Threats'].append("Regulatory changes impacting operations")
        if debt_to_equity and debt_to_equity > 1.5:
            swot['Threats'].append("Interest rate sensitivity due to high debt levels")
    
    except Exception as e:
        # Fallback SWOT analysis
        swot['Strengths'] = ["Established market presence", "Diverse revenue streams"]
        swot['Weaknesses'] = ["Need to improve operational efficiency", "Dependence on current market conditions"]
        swot['Opportunities'] = ["Market expansion possibilities", "Digital transformation opportunities"]
        swot['Threats'] = ["Economic uncertainty", "Increasing competitive pressure"]
    
    return swot

# ----------------------------- PDF REPORT WITH UNICODE SUPPORT -----------------------------

class CombinedPDF(FPDF):
    def __init__(self):
        super().__init__()
        # Agregar soporte para caracteres especiales
        self.add_page()
    
    def header(self): 
        pass
    
    def footer(self): 
        pass
    
    def section_title(self, title):
        self.ln(6)
        self.set_font("Arial", "B", 12)
        # Limpiar caracteres especiales del título
        clean_title = self.clean_text(title)
        self.cell(0, 8, clean_title, ln=True, align="C")
        self.ln(4)
    
    def clean_text(self, text):
        """Clean text for PDF compatibility - remove unsupported characters"""
        if not text:
            return ""
        
        # Reemplazar caracteres problemáticos
        replacements = {
            '€': 'EUR',
            '£': 'GBP',
            '¥': 'JPY',
            '₹': 'INR',
            '§': 'S',
            '©': '(c)',
            '®': '(R)',
            '™': 'TM',
            '°': ' deg',
            '±': '+/-',
            '×': 'x',
            '÷': '/',
            'α': 'alpha',
            'β': 'beta',
            'γ': 'gamma',
            'δ': 'delta'
        }
        
        clean_text = str(text)
        for char, replacement in replacements.items():
            clean_text = clean_text.replace(char, replacement)
        
        # Remover cualquier otro carácter no ASCII
        clean_text = re.sub(r'[^\x00-\x7F]+', ' ', clean_text)
        
        return clean_text
    
    def add_kvk_table(self, df):
        col_widths = [60, 120]
        margin_x = (210 - sum(col_widths)) / 2
        self.set_x(margin_x)
        self.set_font("Arial", "B", 8)
        self.cell(col_widths[0], 8, "Field", border=1, align="C")
        self.cell(col_widths[1], 8, "Value", border=1, align="C")
        self.ln()
        self.set_font("Arial", "", 8)
        for _, row in df.iterrows():
            self.set_x(margin_x)
            field = self.clean_text(str(row["Field"]))
            value = self.clean_text(str(row["Value"]))
            self.cell(col_widths[0], 8, field, border=1)
            self.cell(col_widths[1], 8, value, border=1)
            self.ln()
    
    def add_df_table(self, df, col_widths, headers):
        margin_x = (210 - sum(col_widths)) / 2
        self.set_x(margin_x)
        self.set_font("Arial", "B", 8)
        
        # Limpiar headers
        clean_headers = [self.clean_text(str(h)) for h in headers]
        for i, h in enumerate(clean_headers):
            self.cell(col_widths[i], 8, h, border=1, align="C")
        self.ln()
        
        self.set_font("Arial", "", 8)
        for _, row in df.iterrows():
            self.set_x(margin_x)
            for i, cell in enumerate(row):
                if isinstance(cell, (int, float)):
                    if pd.isna(cell):
                        val = ""
                    else:
                        # Formatear números sin caracteres especiales
                        if abs(cell) >= 1000:
                            val = f"{cell:,.0f}"
                        else:
                            val = f"{cell:.2f}"
                else:
                    val = str(cell)
                
                # Limpiar el texto
                val_clean = self.clean_text(val)
                align = "R" if str(cell).replace(".", "").replace("-", "").replace(",", "").isdigit() else "L"
                self.cell(col_widths[i], 8, val_clean, border=1, align=align)
            self.ln()

def add_news_to_pdf(pdf, articles, company_name):
    """Add news section to PDF report"""
    if articles:
        pdf.section_title("Recent Risk News")
        pdf.set_font("Arial", "B", 10)
        pdf.cell(0, 8, f"Recent news for {company_name}", ln=True)
        pdf.ln(2)
        
        pdf.set_font("Arial", "", 8)
        for i, article in enumerate(articles[:5]):  # Solo primeros 5 en PDF
            title = article.get("title", "No title")
            source = article.get("source", {}).get("name", "Unknown")
            date = article.get("publishedAt", "")[:10]
            
            # Limpiar el texto para PDF
            clean_title = pdf.clean_text(title)
            clean_source = pdf.clean_text(source)
            
            # Truncar título si es muy largo
            if len(clean_title) > 80:
                clean_title = clean_title[:77] + "..."
            
            pdf.multi_cell(0, 4, f"{i+1}. {clean_title}", border=0)
            pdf.cell(0, 4, f"   Source: {clean_source} | Date: {date}", ln=True)
            pdf.ln(2)

def create_download_link(pdf_path, filename):
    """Create download link"""
    try:
        with open(pdf_path, "rb") as f:
            pdf_data = f.read()
        b64_pdf = base64.b64encode(pdf_data).decode()
        return f'<a href="data:application/pdf;base64,{b64_pdf}" download="{filename}" class="download-btn">📄 Download Report</a>'
    except Exception as e:
        st.error(f"Error creating download link: {str(e)}")
        return ""

# ----------------------------- MAIN APPLICATION -----------------------------

def main():
    # Header section
    col1, col2 = st.columns([3, 1])
    with col1:
        st.markdown('<div class="main-header">Financial Analysis Platform</div>', unsafe_allow_html=True)
        st.markdown('<div class="sub-header">Professional Financial Statement Analysis</div>', unsafe_allow_html=True)
    with col2:
        st.image("https://www.accenture.com/content/dam/accenture/final/accenture-com/logo/Accenture-Logo.png", width=150)

    st.markdown("---")

    # Upload section in main area
    st.markdown("### 📁 Document Upload")
    
    col1, col2 = st.columns(2)
    
    with col1:
        st.markdown('<div class="upload-section">', unsafe_allow_html=True)
        kvk_file = st.file_uploader(
            "**Company Registration Document**",
            type="pdf",
            help="Upload KVK/Chamber of Commerce registration PDF"
        )
        st.markdown('</div>', unsafe_allow_html=True)
    
    with col2:
        st.markdown('<div class="upload-section">', unsafe_allow_html=True)
        annual_file = st.file_uploader(
            "**Annual Financial Report**",
            type="pdf", 
            help="Upload company annual report PDF (optional)"
        )
        st.markdown('</div>', unsafe_allow_html=True)

    # Analyze button centered
    col1, col2, col3 = st.columns([1, 2, 1])
    with col2:
        analyze_clicked = st.button("** Analyze Financial Statements**", 
                                  type="primary", 
                                  use_container_width=True,
                                  help="Extract and analyze financial data from uploaded documents")

    if analyze_clicked:
        if not kvk_file:
            st.error("Please upload a Company Registration Document to proceed.")
            return

        with st.spinner("Analyzing financial documents..."):
            # Save temporary KVK file
            with tempfile.NamedTemporaryFile(delete=False, suffix=".pdf") as tmp_kvk:
                tmp_kvk.write(kvk_file.getvalue())
                kvk_path = tmp_kvk.name
            
            try:
                # Process KVK data first
                kvk_df_result = extract_kvk_data_robust(kvk_path)
                
                # Extract company name from KVK for fallback
                company_name = ""
                name_row = kvk_df_result[kvk_df_result['Field'] == 'Name']
                if not name_row.empty and name_row['Value'].iloc[0] != "Not found":
                    company_name = name_row['Value'].iloc[0]
                else:
                    company_name = "Unknown Company"
                
            except Exception as e:
                st.error(f"Error processing company registration: {str(e)}")
                return
            finally:
                try:
                    os.unlink(kvk_path)
                except:
                    pass
            
            # Process financial data - try PDF first, then Yahoo Finance
            financial_data_extracted = False
            data_source = "Unknown"
            ticker = None
            
            if annual_file:
                try:
                    with tempfile.NamedTemporaryFile(delete=False, suffix=".pdf") as tmp_annual:
                        tmp_annual.write(annual_file.getvalue())
                        annual_path = tmp_annual.name
                    
                    doc = fitz.open(annual_path)
                    full_text = "\n".join(page.get_text("text") for page in doc)
                    years = extract_years(full_text)

                    # ⭐⭐ NUEVO: Extraer nombre de la compañía del reporte anual ⭐⭐
                    annual_company_name = extract_company_name_from_annual(full_text)
                    if annual_company_name and annual_company_name != "Unknown Company":
                        company_name = annual_company_name  # Usar el nombre del reporte anual
                        st.sidebar.success(f"Company name from annual report: {company_name}")

                    # Extract data using original PDF methods
                    bc_df = extract_balance_sheet(full_text, years)
                    pl_df = extract_profit_loss(doc, find_profit_loss_page(doc), years)
                    cash_df = extract_cash_flow(doc, find_cash_flow_page(doc), years)
                    
                    # Check if we got meaningful data
                    if not bc_df.empty and not pl_df.empty:
                        # Asegurar que los DataFrames sean numéricos
                        bc_df_numeric = bc_df.copy()
                        pl_df_numeric = pl_df.copy()
                        
                        for year in years:
                            bc_df_numeric[year] = pd.to_numeric(bc_df_numeric[year], errors='coerce')
                            pl_df_numeric[year] = pd.to_numeric(pl_df_numeric[year], errors='coerce')
                        
                        ratios_df = compute_ratios(bc_df_numeric, pl_df_numeric, years).round(3)
                        financial_data_extracted = True
                        data_source = "Annual Report PDF"
                        company_info = None
                    else:
                        financial_data_extracted = False
                    
                    doc.close()
                    os.unlink(annual_path)
                    
                except Exception as e:
                    st.warning(f"PDF extraction failed: {str(e)}")
                    financial_data_extracted = False
            
            # Use Yahoo Finance if PDF failed or wasn't provided
            if not financial_data_extracted and company_name:
                ticker = find_company_ticker(company_name)
                
                if ticker:
                    bc_df, pl_df, cash_df, company_info, years = get_yahoo_finance_data(ticker, company_name)
                    
                    if bc_df is not None and pl_df is not None:
                        ratios_df = compute_ratios(bc_df, pl_df, years).round(3)
                        financial_data_extracted = True
                        data_source = f"Yahoo Finance ({ticker})"
            
            if not financial_data_extracted:
                st.error("Could not extract financial data. Please check your documents and try again.")
                return

            # Asegurar que todos los DataFrames sean numéricos para las métricas
            bc_df_numeric = bc_df.copy()
            pl_df_numeric = pl_df.copy()
            
            for year in years:
                bc_df_numeric[year] = pd.to_numeric(bc_df_numeric[year], errors='coerce')
                pl_df_numeric[year] = pd.to_numeric(pl_df_numeric[year], errors='coerce')

            # Generate charts and SWOT analysis
            charts = create_financial_charts(bc_df_numeric, pl_df_numeric, cash_df, ratios_df, years)
            swot_analysis = generate_swot_analysis(bc_df_numeric, pl_df_numeric, ratios_df, years, company_name)
            
            # Get risk news - MEJORADO
            with st.spinner("📰 Fetching relevant news..."):
                news_api_key = "e35e621ac9a245a8879ef9a2505f2646"
                news_articles = get_risk_news(company_name, news_api_key)
                
                # Si no hay resultados, intentar con el ticker
                if not news_articles and ticker:
                    news_articles = get_risk_news(ticker, news_api_key)

            # Display success message
            st.markdown(f'<div class="success-box"><strong>✅ Analysis Complete</strong><br>Data source: {data_source}</div>', unsafe_allow_html=True)

            # Results section
            st.markdown("## 📊 Financial Analysis Results")
            
            # Key metrics at the top
            col1, col2, col3, col4 = st.columns(4)
            with col1:
                st.markdown('<div class="metric-card">', unsafe_allow_html=True)
                revenue = pl_df_numeric[pl_df_numeric['Items'] == 'Revenue']
                if not revenue.empty:
                    latest_rev = revenue[years[0]].iloc[0] if pd.notna(revenue[years[0]].iloc[0]) else revenue[years[1]].iloc[0]
                    if pd.notna(latest_rev):
                        st.metric("Revenue", f"${latest_rev:,.0f}")
                    else:
                        st.metric("Revenue", "N/A")
                st.markdown('</div>', unsafe_allow_html=True)
            
            with col2:
                st.markdown('<div class="metric-card">', unsafe_allow_html=True)
                profit = pl_df_numeric[pl_df_numeric['Items'] == 'Profit for the year']
                if not profit.empty:
                    latest_profit = profit[years[0]].iloc[0] if pd.notna(profit[years[0]].iloc[0]) else profit[years[1]].iloc[0]
                    if pd.notna(latest_profit):
                        st.metric("Net Profit", f"${latest_profit:,.0f}")
                    else:
                        st.metric("Net Profit", "N/A")
                st.markdown('</div>', unsafe_allow_html=True)
            
            with col3:
                st.markdown('<div class="metric-card">', unsafe_allow_html=True)
                assets = bc_df_numeric[bc_df_numeric['Items'] == 'Total Assets']
                if not assets.empty:
                    latest_assets = assets[years[0]].iloc[0] if pd.notna(assets[years[0]].iloc[0]) else assets[years[1]].iloc[0]
                    if pd.notna(latest_assets):
                        st.metric("Total Assets", f"${latest_assets:,.0f}")
                    else:
                        st.metric("Total Assets", "N/A")
                st.markdown('</div>', unsafe_allow_html=True)
            
            with col4:
                st.markdown('<div class="metric-card">', unsafe_allow_html=True)
                current_ratio = ratios_df[ratios_df['Ratios'] == 'Current Ratio']
                if not current_ratio.empty:
                    latest_ratio = current_ratio[years[0]].iloc[0] if pd.notna(current_ratio[years[0]].iloc[0]) else current_ratio[years[1]].iloc[0]
                    if pd.notna(latest_ratio):
                        st.metric("Current Ratio", f"{latest_ratio:.2f}")
                    else:
                        st.metric("Current Ratio", "N/A")
                st.markdown('</div>', unsafe_allow_html=True)

            # Tabs for detailed analysis - INCLUYENDO PESTAÑA DE NOTICIAS
            tab1, tab2, tab3, tab4, tab5, tab6, tab7 = st.tabs([
                "🏢 Company Info", 
                "💰 Balance Sheet", 
                "📈 Profit & Loss", 
                "💸 Financial Ratios",
                "📊 Financial Charts",
                "🎯 SWOT Analysis",
                "📰 Risk News"
            ])

            with tab1:
                col1, col2 = st.columns([2, 1])
                with col1:
                    st.subheader("Company Registration Details")
                    st.dataframe(kvk_df_result, use_container_width=True, hide_index=True)
                
                with col2:
                    if company_info is not None:
                        st.subheader("Market Information")
                        st.dataframe(company_info, use_container_width=True, hide_index=True)

            with tab2:
                st.subheader("Balance Sheet Analysis")
                if "Yahoo Finance" in data_source:
                    st.dataframe(bc_df.style.format({
                        years[0]: '{:,.0f}',
                        years[1]: '{:,.0f}'
                    }), use_container_width=True, hide_index=True)
                else:
                    st.dataframe(bc_df, use_container_width=True, hide_index=True)

            with tab3:
                st.subheader("Profit & Loss Statement")
                if "Yahoo Finance" in data_source:
                    st.dataframe(pl_df.style.format({
                        years[0]: '{:,.0f}',
                        years[1]: '{:,.0f}'
                    }), use_container_width=True, hide_index=True)
                else:
                    st.dataframe(pl_df, use_container_width=True, hide_index=True)

            with tab4:
                st.subheader("Financial Ratios")
                st.dataframe(ratios_df, use_container_width=True, hide_index=True)

            with tab5:
                st.subheader("Financial Performance Charts")
                
                if 'revenue_profit' in charts:
                    st.plotly_chart(charts['revenue_profit'], use_container_width=True)
                
                col1, col2 = st.columns(2)
                
                with col1:
                    if 'balance_sheet' in charts:
                        st.plotly_chart(charts['balance_sheet'], use_container_width=True)
                    if 'profitability' in charts:
                        st.plotly_chart(charts['profitability'], use_container_width=True)
                
                with col2:
                    if 'liquidity' in charts:
                        st.plotly_chart(charts['liquidity'], use_container_width=True)

            with tab6:
                st.subheader("SWOT Analysis")
                
                col1, col2 = st.columns(2)
                
                with col1:
                    st.markdown('<div class="swot-box swot-strengths">', unsafe_allow_html=True)
                    st.markdown("### 💪 Strengths")
                    for strength in swot_analysis['Strengths']:
                        st.markdown(f"• {strength}")
                    st.markdown('</div>', unsafe_allow_html=True)
                    
                    st.markdown('<div class="swot-box swot-weaknesses">', unsafe_allow_html=True)
                    st.markdown("### 📉 Weaknesses")
                    for weakness in swot_analysis['Weaknesses']:
                        st.markdown(f"• {weakness}")
                    st.markdown('</div>', unsafe_allow_html=True)
                
                with col2:
                    st.markdown('<div class="swot-box swot-opportunities">', unsafe_allow_html=True)
                    st.markdown("### 🚀 Opportunities")
                    for opportunity in swot_analysis['Opportunities']:
                        st.markdown(f"• {opportunity}")
                    st.markdown('</div>', unsafe_allow_html=True)
                    
                    st.markdown('<div class="swot-box swot-threats">', unsafe_allow_html=True)
                    st.markdown("### ⚠️ Threats")
                    for threat in swot_analysis['Threats']:
                        st.markdown(f"• {threat}")
                    st.markdown('</div>', unsafe_allow_html=True)

            with tab7:
                st.subheader("📰 Latest Company News & Risk Alerts")
                
                # Agregar opción para buscar noticias manualmente
                col1, col2 = st.columns([3, 1])
                with col1:
                    search_query = st.text_input("Search for specific news:", 
                                               value=f"{company_name} financial risk",
                                               key="news_search_input")
                with col2:
                    if st.button("🔍 Search News", key="news_search_btn"):
                        with st.spinner("Searching..."):
                            news_articles = get_risk_news(search_query, news_api_key)
                            st.rerun()
                
                display_news_articles(news_articles, company_name)
                
                # Información sobre la fuente de noticias
                with st.expander("ℹ️ About News Source"):
                    st.write("""
                    **News Source**: NewsAPI.org
                    **Coverage**: Global news sources
                    **Timeframe**: Last 30 days
                    **Focus**: Financial, regulatory, and risk-related news
                    
                    *Note: News availability depends on media coverage and API limits.*
                    """)

            # Report generation
            st.markdown("---")
            st.markdown("## 📄 Generate Professional Report")
            
            with st.spinner("Preparing final report..."):
                try:
                    pdf = CombinedPDF()
                    pdf.add_page()
                    
                    pdf.section_title("Financial Analysis Report")
                    pdf.ln(5)
                    
                    # Company Information
                    pdf.section_title("Company Information")
                    pdf.add_kvk_table(kvk_df_result)
                    
                    # Yahoo Finance Info if available
                    if company_info is not None:
                        pdf.section_title("Market Information")
                        pdf.add_kvk_table(company_info)
                    
                    # Balance Sheet
                    if not bc_df.empty:
                        pdf.section_title("Balance Sheet")
                        pdf.add_df_table(bc_df[["Items", *years]], [70, 40, 40], headers=["Items", *years])
                    
                    # Profit & Loss Statement
                    if not pl_df.empty:
                        pdf.section_title("Profit & Loss Statement")
                        pdf.add_df_table(pl_df[["Items", *years]], [70, 40, 40], headers=["Items", *years])
                    
                    # Financial Ratios
                    if not ratios_df.empty:
                        pdf.section_title("Financial Ratios")
                        pdf.add_df_table(ratios_df[["Ratios", *years]], [70, 40, 40], headers=["Ratios", *years])
                    
                    # Cash Flow Statement
                    if not cash_df.empty:
                        pdf.section_title("Cash Flow Statement")
                        pdf.add_df_table(cash_df[["Items", *years]], [80, 35, 35], headers=["Items", *years])
                    
                    # News Section
                    if news_articles:
                        add_news_to_pdf(pdf, news_articles, company_name)

                    pdf_output_path = "Financial_Analysis_Report.pdf"
                    pdf.output(pdf_output_path)

                    st.markdown("### Download Final Report")
                    st.markdown(create_download_link(pdf_output_path, "Financial_Analysis_Report.pdf"), unsafe_allow_html=True)
                    
                    st.success("✅ PDF report generated successfully!")
                    
                except Exception as pdf_error:
                    st.error(f"Error generating PDF report: {str(pdf_error)}")
                    st.info("💡 Tip: This error usually occurs due to special characters in the data. The system has been updated to handle this automatically.")

if __name__ == "__main__":
    main()
