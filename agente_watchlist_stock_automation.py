"""
Agente Watchlist - Applicazione Streamlit Avanzata
Lingua: Italiano
Descrizione:
  Web-app Streamlit che fornisce:
   - Watchlist dinamica e autosufficiente per analisi quantitativa + qualitativa
   - Screening automatico per titoli "hot" (momentum + score)
   - Estrazione automatica di informazioni qualitative (Yahoo Finance + sintesi)
   - Analisi news daily con semplice sentiment e previsione "up/down/neutral" con motivazione testuale
   - Sezione strategie opzioni con payoff, backtest storico e simulatore
   - Generazione report Excel (Qualitativo, Finanziario, Checks, Options, News, Backtest)
   - Modalità di aggiornamento automatico tramite script CLI separato (cron / GitHub Actions)
   - Notifiche (Telegram) opzionali: invio alert quotidiani per titoli hot (configurare TOKEN & CHAT_ID)

Requisiti (pip):
  pip install streamlit yfinance pandas numpy openpyxl xlsxwriter requests beautifulsoup4 matplotlib python-dotenv

Come avviare in locale:
  1) Salva questo file come `app.py` nella cartella del progetto
  2) Crea una cartella `reports/`
  3) (Opzionale) crea file `.env` con TELEGRAM_TOKEN e TELEGRAM_CHAT_ID se vuoi notifiche
  4) Installa dipendenze: pip install -r requirements.txt  (vedi README)
  5) Avvia: streamlit run app.py

Aggiornamento automatico / schedulazione (opzioni):
  - Cron (Linux/macOS): scrivi una job che esegua `python updater.py` ogni giorno alle 07:00
  - GitHub Actions: crea workflow con schedule cron che esegue lo script `updater.py` e commit dei report
  - Streamlit Cloud: puoi deployare l'app ma per job schedulati usa GitHub Actions

Note su limitazioni:
  - Lo scraping di Yahoo Finance è leggero e non garantito; per produzione usa API ufficiali.
  - Le previsioni basate su sentiment leggero sono indicazioni, NON raccomandazioni di trading.

"""

import os
import time
import datetime as dt
import io
import math
import requests
from bs4 import BeautifulSoup
from dotenv import load_dotenv

import pandas as pd
import numpy as np
import yfinance as yf
import streamlit as st
import matplotlib.pyplot as plt

# Carica .env se presente
load_dotenv()
TELEGRAM_TOKEN = os.getenv('TELEGRAM_TOKEN')
TELEGRAM_CHAT_ID = os.getenv('TELEGRAM_CHAT_ID')

# ---------- CONFIG ----------
REPORTS_DIR = 'reports'
os.makedirs(REPORTS_DIR, exist_ok=True)

# Default watchlist (modificabile dall'utente nell'app)
DEFAULT_WATCHLIST = ['AAPL','MSFT','TSLA','NVDA','AMZN']

# Pesi per scoring
WEIGHTS = {
    'financial_health': 30,
    'growth_prospects': 20,
    'competitive_advantage': 15,
    'innovation_strategy': 10,
    'governance': 10,
    'risk_profile': 15,
}

# Semplice lessico per sentiment (estendibile)
POS_WORDS = set(['gain','upgrade','beat','record','strong','growth','surge','outperform','positive','increase'])
NEG_WORDS = set(['miss','downgrade','drop','weak','loss','decline','recall','lawsuit','investigation','negative'])

# ---------- UTILS ----------

def safe_div(a, b):
    try:
        return a / b if b else np.nan
    except Exception:
        return np.nan

# Invio notifiche Telegram opzionale
def send_telegram(text):
    token = TELEGRAM_TOKEN
    chat = TELEGRAM_CHAT_ID
    if not token or not chat:
        return False
    url = f'https://api.telegram.org/bot{token}/sendMessage'
    try:
        requests.post(url, json={'chat_id': chat, 'text': text})
        return True
    except Exception:
        return False

# ---------- FETCH QUALITATIVE ----------

def fetch_qualitative_yahoo(ticker):
    url = f'https://finance.yahoo.com/quote/{ticker}/profile'
    headers = {'User-Agent': 'Mozilla/5.0'}
    try:
        r = requests.get(url, headers=headers, timeout=8)
        soup = BeautifulSoup(r.text, 'html.parser')
        profile_section = soup.find('section', {'data-test':'qsp-profile'})
        description = ''
        if profile_section:
            p = profile_section.find('p')
            description = p.get_text(separator=' ').strip() if p else ''
        # trova sector/industry
        info = soup.find_all('p', class_='D(ib) Va(t)')
        extras = ' | '.join([x.get_text(separator=' ').strip() for x in info]) if info else ''
        return {
            'Business Model': description or f'Descrizione non trovata per {ticker}.',
            'Supply Chain': 'Supply chain: da verificare manualmente (dati non strutturati).',
            'Market & Competition': extras,
            'Strategy': '',
            'Governance': '',
            'Risks': ''
        }
    except Exception:
        return {k: '' for k in ['Business Model','Supply Chain','Market & Competition','Strategy','Governance','Risks']}

# ---------- QUANT METRICS ----------

def compute_quant_metrics(ticker_obj):
    info = ticker_obj.info
    res = {}
    try:
        financials = ticker_obj.financials.T
        if 'Total Revenue' in financials.columns:
            revenue = financials['Total Revenue'].dropna().astype(float)
        elif 'Revenue' in financials.columns:
            revenue = financials['Revenue'].dropna().astype(float)
        else:
            revenue = pd.Series()
        sales_growth = np.nan
        if len(revenue) >= 2:
            sales_growth = (revenue.iloc[0] - revenue.iloc[-1]) / abs(revenue.iloc[-1])
    except Exception:
        sales_growth = np.nan

    try:
        cf = ticker_obj.cashflow.T
        if 'Free Cash Flow' in cf.columns:
            fcf = cf['Free Cash Flow'].dropna().astype(float).iloc[0]
        else:
            fcf = (cf.get('Total Cash From Operating Activities', pd.Series()).dropna().astype(float).iloc[0]
                   - cf.get('Capital Expenditures', pd.Series()).dropna().astype(float).iloc[0])
    except Exception:
        fcf = np.nan

    try:
        bs = ticker_obj.balance_sheet.T
        total_debt = bs.get('Long Term Debt', pd.Series()).dropna().astype(float).iloc[0] if 'Long Term Debt' in bs.columns else np.nan
        total_equity = bs.get('Total Stockholder Equity', pd.Series()).dropna().astype(float).iloc[0] if 'Total Stockholder Equity' in bs.columns else np.nan
        debt_to_equity = safe_div(total_debt, total_equity) * 100
    except Exception:
        debt_to_equity = np.nan

    marg = info.get('operatingMargins', np.nan)
    pe = info.get('forwardPE') or info.get('trailingPE') or np.nan

    res.update({
        'sales_growth': float(sales_growth) if not pd.isna(sales_growth) else np.nan,
        'fcf': float(fcf) if not pd.isna(fcf) else np.nan,
        'debt_to_equity_pct': float(debt_to_equity) if not pd.isna(debt_to_equity) else np.nan,
        'operating_margin': float(marg) if not pd.isna(marg) else np.nan,
        'pe': float(pe) if not pd.isna(pe) else np.nan,
    })
    return res

# ---------- YES/NO CHECKS ----------

def yes_no_checks(metrics, info):
    checks = {}
    checks['Sales growth stable?'] = bool(not pd.isna(metrics.get('sales_growth')) and metrics['sales_growth'] > 0.05)
    checks['Earnings trend positive?'] = bool(info.get('earningsQuarterlyGrowth', 0) > 0)
    checks['FCF positivo?'] = bool(not pd.isna(metrics.get('fcf')) and metrics['fcf'] > 0)
    checks['Operating margin > 0?'] = bool(metrics.get('operating_margin') and metrics['operating_margin'] > 0)
    checks['Debt-to-Equity < 60%?'] = bool(not pd.isna(metrics.get('debt_to_equity_pct')) and metrics['debt_to_equity_pct'] < 60)
    checks['Dividendi presenti?'] = bool(info.get('dividendYield'))
    checks['PE < 30?'] = bool(metrics.get('pe') and metrics.get('pe') < 30)
    return checks

# ---------- SCORING ----------

def compute_score(metrics, qualitative_scores):
    total = 0
    for cat, w in WEIGHTS.items():
        s = qualitative_scores.get(cat, 5)
        total += s * (w / 10.0)
    score = safe_div(total, sum(WEIGHTS.values())) * 10
    score = max(1, min(10, round(score, 2)))
    return score

# ---------- NEWS + SENTIMENT SIMPLE ----------

def fetch_news_yahoo(ticker, n=5):
    url = f'https://finance.yahoo.com/quote/{ticker}/news'
    headers = {'User-Agent': 'Mozilla/5.0'}
    try:
        r = requests.get(url, headers=headers, timeout=8)
        soup = BeautifulSoup(r.text, 'html.parser')
        items = []
        news_cards = soup.find_all('li', attrs={'class': 'js-stream-content'})
        for card in news_cards[:n]:
            title = card.find('h3')
            if title:
                text = title.get_text(separator=' ').strip()
                link_tag = card.find('a', href=True)
                link = 'https://finance.yahoo.com' + link_tag['href'] if link_tag else ''
                items.append({'title': text, 'link': link})
        return items
    except Exception:
        return []


def sentiment_simple(text):
    t = text.lower()
    pos = sum(1 for w in POS_WORDS if w in t)
    neg = sum(1 for w in NEG_WORDS if w in t)
    score = pos - neg
    if score > 0:
        return 'positive', score
    if score < 0:
        return 'negative', score
    return 'neutral', 0

# ---------- HOT PICKS: screening ----------

def screen_hot_picks(watchers):
    rows = []
    for tk in watchers:
        try:
            t = yf.Ticker(tk)
            m = compute_quant_metrics(t)
            hist = t.history(period='7d')
            if hist.empty:
                momentum = 0
            else:
                momentum = (hist['Close'][-1] - hist['Close'][0]) / hist['Close'][0]
            # simple hotness metric: momentum * (1 if PE<30 else 0.7)
            pe = m.get('pe') or 999
            hot = momentum * (1 if pe < 30 else 0.7)
            rows.append({'ticker':tk, 'momentum7d':momentum, 'pe':pe, 'hot_metric':hot})
        except Exception:
            rows.append({'ticker':tk, 'momentum7d':np.nan, 'pe':np.nan, 'hot_metric':np.nan})
    df = pd.DataFrame(rows).sort_values('hot_metric', ascending=False)
    return df

# ---------- OPTIONS PAYOFF & BACKTEST ----------

def payoff_long_call(strike, premium, spot_range):
    return np.maximum(spot_range - strike, 0) - premium


def payoff_covered_call(spot_range, stock_price, strike, premium_received):
    stock_payoff = spot_range - stock_price
    call_payoff = np.minimum(spot_range - strike, 0) + premium_received
    return stock_payoff + call_payoff


def payoff_protective_put(spot_range, stock_price, strike, premium_paid):
    stock_payoff = spot_range - stock_price
    put_payoff = np.maximum(strike - spot_range, 0) - premium_paid
    return stock_payoff + put_payoff


def backtest_option_strategy(ticker, strategy, strike, premium, lookback_years=2, expiry_days=30):
    # semplice backtest: per ogni giorno in lookback, prendi prezzo S0, cerca prezzo a expiry_days e calcola payoff
    t = yf.Ticker(ticker)
    hist = t.history(period=f'{lookback_years}y')
    hist = hist['Close'].dropna()
    results = []
    for i in range(len(hist)-expiry_days):
        S0 = hist.iloc[i]
        Sexp = hist.iloc[i+expiry_days]
        if strategy == 'Long Call':
            payoff = max(Sexp - strike, 0) - premium
        elif strategy == 'Covered Call':
            payoff = (Sexp - S0) + min(Sexp - strike, 0) + premium
        elif strategy == 'Protective Put':
            payoff = (Sexp - S0) + max(strike - Sexp, 0) - premium
        else:
            payoff = 0
        results.append(payoff)
    if results:
        arr = np.array(results)
        return {'avg': float(np.mean(arr)), 'win_rate': float((arr>0).mean()), 'trades': len(arr)}
    return {'avg': np.nan, 'win_rate': np.nan, 'trades': 0}

# ---------- REPORT GENERATOR ----------

def generate_excel_report(ticker, qual, metrics, checks, score, news_items=None, options_df=None, backtest_res=None):
    fname = os.path.join(REPORTS_DIR, f"{ticker}_report_{dt.date.today().isoformat()}.xlsx")
    with pd.ExcelWriter(fname, engine='xlsxwriter') as writer:
        pd.DataFrame(list(qual.items()), columns=['Sezione','Contenuto']).to_excel(writer, sheet_name='Qualitativo', index=False)
        pd.DataFrame(list(metrics.items()), columns=['Metric','Value']).to_excel(writer, sheet_name='Finanziario', index=False)
        pd.DataFrame(list(checks.items()), columns=['Check','Result']).to_excel(writer, sheet_name='Checks', index=False)
        pd.DataFrame({'Ticker':[ticker],'GlobalScore':[score]}).to_excel(writer, sheet_name='Summary', index=False)
        if news_items:
            pd.DataFrame(news_items).to_excel(writer, sheet_name='News', index=False)
        if options_df is not None and not options_df.empty:
            options_df.to_excel(writer, sheet_name='Options', index=False)
        if backtest_res:
            pd.DataFrame([backtest_res]).to_excel(writer, sheet_name='Backtest', index=False)
    return fname

# ---------- STREAMLIT UI ----------

st.set_page_config(page_title='Agente Watchlist Avanzato', layout='wide')
st.title('Agente Watchlist — Analisi Automatica & Strategie Opzioni')
st.markdown('Usa questa app come base: è progettata per essere estesa e messa in produzione con job schedulati.')

# Sidebar: impostazioni globali
st.sidebar.header('Impostazioni')
watchlist_text = st.sidebar.text_area('Watchlist (virgola separati)', value=','.join(DEFAULT_WATCHLIST), height=120)
watchers = [w.strip().upper() for w in watchlist_text.split(',') if w.strip()]
refresh_button = st.sidebar.button('Esegui scansione ora')
auto_notify = st.sidebar.checkbox('Invia notifiche Telegram per hot picks', value=False)

st.sidebar.markdown('---')
st.sidebar.markdown('Schedulazione automatica: vedi README per esempi di cron / GitHub Actions.')

# Main: scansione e summary
if refresh_button:
    st.info('Esecuzione scansione...')
    summary_rows = []
    hot_df = screen_hot_picks(watchers)

    for tk in watchers:
        with st.spinner(f'Analizzo {tk}...'):
            t = yf.Ticker(tk)
            info = t.info
            metrics = compute_quant_metrics(t)
            checks = yes_no_checks(metrics, info)
            qual = fetch_qualitative_yahoo(tk)
            # qualitative scoring: lasciamo all'utente sliders per modificare; qui default neutral
            qualitative_scores = {
                'financial_health': 7 if (metrics.get('debt_to_equity_pct') is not None and metrics.get('debt_to_equity_pct') < 60) else 4,
                'growth_prospects': 6 if (metrics.get('sales_growth') and metrics.get('sales_growth')>0.05) else 4,
                'competitive_advantage': 5,
                'innovation_strategy': 5,
                'governance': 5,
                'risk_profile': 6 if (metrics.get('fcf') and metrics.get('fcf')>0) else 4,
            }
            score = compute_score(metrics, qualitative_scores)
            # news + sentiment
            news_items = fetch_news_yahoo(tk, n=5)
            sentiments = []
            for ni in news_items:
                s, sc = sentiment_simple(ni['title'])
                ni['sentiment'] = s
                ni['sentiment_score'] = sc
                sentiments.append(sc)
            news_sentiment_agg = np.nansum(sentiments) if sentiments else 0
            # simple prediction rule
            momentum7d = 0
            try:
                hist = t.history(period='7d')
                if not hist.empty:
                    momentum7d = float((hist['Close'][-1] - hist['Close'][0]) / hist['Close'][0])
            except Exception:
                momentum7d = 0
            prediction = 'neutral'
            rationale = []
            if news_sentiment_agg > 0 and momentum7d > 0:
                prediction = 'up'
                rationale.append('Sentiment news positivo e momentum positivo')
            elif news_sentiment_agg < 0 and momentum7d < 0:
                prediction = 'down'
                rationale.append('Sentiment news negativo e momentum negativo')
            else:
                if news_sentiment_agg > 0:
                    rationale.append('Sentiment news positivo')
                elif news_sentiment_agg < 0:
                    rationale.append('Sentiment news negativo')
                if momentum7d > 0:
                    rationale.append('Momentum 7d positivo')

            # backtest esempio per Covered Call 30d
            backtest_res = backtest_option_strategy(tk, 'Covered Call', strike=math.floor(info.get('regularMarketPrice', 100)), premium=2, lookback_years=1, expiry_days=30)

            # save report
            excel = generate_excel_report(tk, qual, metrics, checks, score, news_items=news_items, options_df=None, backtest_res=backtest_res)

            summary_rows.append({'Ticker':tk, 'Score':score, 'Prediction':prediction, 'Rationale':'; '.join(rationale)})

    hot_picks = hot_df.head(10)
    st.success('Scansione completata — riepilogo')
    st.table(pd.DataFrame(summary_rows))
    st.markdown('### Hot picks (momentum + valutazione)')
    st.dataframe(hot_picks)

    if auto_notify and not hot_picks.empty:
        text = 'Hot picks oggi:
' + '
'.join(hot_picks['ticker'].tolist())
        send_telegram(text)
        st.info('Notifica inviata (Telegram)')

# Interactive ticker explorer
st.markdown('---')
colL, colR = st.columns([2,1])
with colL:
    st.subheader('Esplora singolo ticker')
    sel = st.text_input('Inserisci ticker da esplorare', value='AAPL').strip().upper()
    if sel:
        t = yf.Ticker(sel)
        info = t.info
        metrics = compute_quant_metrics(t)
        checks = yes_no_checks(metrics, info)
        qual = fetch_qualitative_yahoo(sel)

        st.markdown(f'### {sel} — Schema rapido')
        st.write('**Business Model (auto):**')
        st.write(qual.get('Business Model'))
        st.write('**Metriche**')
        st.write(metrics)
        st.write('**Checks**')
        st.write(checks)

        st.write('**Grafico storico 1y**')
        hist = t.history(period='1y')
        st.line_chart(hist['Close'])

        st.write('**News recenti (estratte)**')
        news_items = fetch_news_yahoo(sel, n=5)
        for ni in news_items:
            st.write('-', ni.get('title'))
            st.write(ni.get('link'))

        st.write('**Simulatore Strategie Opzioni (payoff)**')
        strat = st.selectbox('Strategia', ['Long Call','Covered Call','Protective Put','Straddle'])
        strike = st.number_input('Strike', value=float(info.get('regularMarketPrice', 100)))
        premium_c = st.number_input('Premio Call', value=2.0)
        premium_p = st.number_input('Premio Put', value=2.0)
        low = max(0, info.get('regularMarketPrice', 100) * 0.5)
        high = info.get('regularMarketPrice', 100) * 1.5
        spot_range = np.linspace(low, high, 200)
        fig, ax = plt.subplots()
        if strat == 'Long Call':
            ax.plot(spot_range, payoff_long_call(strike, premium_c, spot_range))
        elif strat == 'Covered Call':
            ax.plot(spot_range, payoff_covered_call(spot_range, info.get('regularMarketPrice',100), strike, premium_c))
        elif strat == 'Protective Put':
            ax.plot(spot_range, payoff_protective_put(spot_range, info.get('regularMarketPrice',100), strike, premium_p))
        else:
            ax.plot(spot_range, payoff_long_call(strike, premium_c, spot_range) + np.maximum(strike - spot_range, 0) - premium_p)
        ax.axhline(0, color='k', linewidth=0.7)
        st.pyplot(fig)

with colR:
    st.subheader('Quick Actions')
    if st.button('Genera report per tutti i watchlist'):
        st.info('Generazione report (potrebbe richiedere tempo)...')
        for tk in watchers:
            t = yf.Ticker(tk)
            info = t.info
            metrics = compute_quant_metrics(t)
            checks = yes_no_checks(metrics, info)
            qual = fetch_qualitative_yahoo(tk)
            excel = generate_excel_report(tk, qual, metrics, checks, score=0)
        st.success('Report generati in reports/')

    st.markdown('### Scarica esempio report')
    sample_files = [f for f in os.listdir(REPORTS_DIR) if f.endswith('.xlsx')]
    if sample_files:
        self = st.selectbox('Seleziona report', sample_files)
        with open(os.path.join(REPORTS_DIR, self), 'rb') as fh:
            st.download_button('Download', fh, file_name=self)
    else:
        st.write('Nessun report disponibile. Esegui prima una scansione.')

st.markdown('---')
st.caption('Questa app fornisce funzioni base e strumenti per diventare un sistema di monitoraggio automatizzato: per produzione consigliamo di integrare provider dati ufficiali, caching e job schedulati.')

# ---------- CLI Updater (script separato: updater.py) ----------
UPDATER_SCRIPT = """
# Salva come updater.py e configura cron o GitHub Actions per eseguirlo quotidianamente
import os
from app import generate_excel_report, fetch_qualitative_yahoo, compute_quant_metrics, yes_no_checks
import yfinance as yf
watchers = %s
for tk in watchers:
    t = yf.Ticker(tk)
    metrics = compute_quant_metrics(t)
    checks = yes_no_checks(metrics, t.info)
    qual = fetch_qualitative_yahoo(tk)
    generate_excel_report(tk, qual, metrics, checks, score=0)
""" % (DEFAULT_WATCHLIST,)

with open('updater.py', 'w') as f:
    f.write(UPDATER_SCRIPT)

# README rapido
README = '''
1) Installare dipendenze:
   pip install streamlit yfinance pandas numpy openpyxl xlsxwriter requests beautifulsoup4 python-dotenv

2) Avviare app in locale:
   streamlit run app.py

3) Per schedulazione giornaliera creare job cron che esegue:
   python updater.py

4) Per notifiche Telegram creare bot e impostare TELEGRAM_TOKEN e TELEGRAM_CHAT_ID nel file .env
'''
with open('README.md', 'w') as f:
    f.write(README)

st.success('Setup iniziale: updater.py e README.md creati nella cartella di progetto.')
