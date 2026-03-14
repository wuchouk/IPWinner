#!/usr/bin/env python3
"""
V3 Email Processor - 下載區批量處理
Parse, classify, and rename all .eml files from 下載區 (488 emails).
Output: renamed .eml copies + classification log JSON.
"""

import email
import email.header
import email.utils
import os
import re
import json
import shutil
import sys
from pathlib import Path
from email import policy
from datetime import datetime
from collections import defaultdict, Counter

# ============================================================
# === Configuration ===
# ============================================================
INPUT_DIR = "/sessions/gifted-awesome-albattani/mnt/Email Processor v2/下載區"
OUTPUT_DIR = "/sessions/gifted-awesome-albattani/mnt/Email Processor v2/V3-下載區-分類結果v4"
LOG_PATH = "/sessions/gifted-awesome-albattani/mnt/Email Processor v2/V3-下載區-分類日誌-v4.json"

# ============================================================
# === Case Number Regex ===
# ============================================================
CASE_RE = re.compile(r'([A-Z0-9]{4}\d{5}[PMDTABCW][A-Z]{2}\d*)')
CASE_BASE_RE = re.compile(r'([A-Z]{4}\d{5})')

# ============================================================
# === Sender Role Database ===
# ============================================================
ROLE_MAP_DOMAIN = {
    # Clients (C)
    'botrista.io': 'C', 'botrista.com': 'C',
    'nippon-alloy.com': 'C',   # 日鐵 NAC
    'koicafe.com': 'C',        # KOI Café
    'realtek.com': 'C',        # Realtek
    'keyxentic.com': 'C',      # KeyXentic
    'tronfuture.com': 'C',     # TRON Future
    'greenhope.com.tw': 'C',   # 綠能
    'anmo.com.tw': 'C',        # AnMo (DinoCapture)
    'octahub.com': 'C',        # 章魚新零售
    'e-suntech.com': 'C',      # E-Sun Tech
    # Agents (A)
    'bskb.com': 'A',
    'cohorizon.com': 'A',
    'chwiplaw.com': 'A',
    'tilleke.com': 'A',
    'shigapatent.com': 'A',
    'twobirds.com': 'A',
    'kimchang.com': 'A',
    'baudelio.com.mx': 'A',
    'atmac.ca': 'A',
    'atmac.com.au': 'A',
    'federislaw.com.ph': 'A',
    'pizzeys.com': 'A',
    'metida.com': 'A',
    'naipc.com': 'A',
    'rfrip.com': 'A',
    'thefirstpatent.co.kr': 'A',
    'intelpat.co.kr': 'A',
    'bakermckenzie.com': 'A',
    'ycmclaw.com': 'A',
    'bhp.com.tw': 'A',
    'taiapatent.com.tw': 'A',
    'wispro.com': 'A',
    'adr-law.com': 'A',
    'markify.com': 'A',
    'smartbiggar.ca': 'A',     # Smart & Biggar - CA
    'altitude-ip.com': 'A',    # Altitude IP
    'boehmert.de': 'A',        # Boehmert - DE
    'tsailee.com.tw': 'A',     # 蔡李法律事務所
    'kspat.com': 'A',          # KASAN - KR
    'clarivate.com': 'A',      # Clarivate (TM watch)
    'cpaglobal.com': 'A',      # CPA Global
    'cpahkltd.com': 'A',       # CPA HK
    'vjp.de': 'A',             # VJP - DE
    'naipo.com': 'A',          # 北美智權
    'webmail.mozlen.com': 'A', # 摩知輪 (TM monitoring)
    'za3c.com.tw': 'A',        # 法律事務所
    'wglaw.com.tw': 'A',       # 法律事務所
    'taiwanlaw.com': 'A',      # 法律事務所
    # Government (G)
    'tiponet.tipo.gov.tw': 'G',
    'tipo.gov.tw': 'G',
    'gov.tw': 'G',
    'google.com': 'G',         # Google Calendar notifications (系統通知)
    # IP Winner (self)
    'ipwinner.com': 'SELF',
    'ipwinner.com.tw': 'SELF',
}

AGENT_CODE_MAP = {
    'bskb.com': 'BSKB',
    'cohorizon.com': 'CHIP',
    'chwiplaw.com': 'CHIP',
    'atmac.ca': 'ATMAC',
    'atmac.com.au': 'ATMAC',
    'naipc.com': 'NAIPC',
    'markify.com': 'Markify',
    'baudelio.com.mx': 'Baudelio',
    'tilleke.com': 'Tilleke',
    'shigapatent.com': 'Shiga',
    'twobirds.com': 'BirdBird',
    'kimchang.com': 'KimChang',
    'pizzeys.com': 'Pizzeys',
    'metida.com': 'Metida',
    'federislaw.com.ph': 'Federis',
    'thefirstpatent.co.kr': 'FirstPatent',
    'intelpat.co.kr': 'Intelpat',
    'bakermckenzie.com': 'BakerMcKenzie',
    'ycmclaw.com': 'YCMC',
    'bhp.com.tw': 'BHP',
    'taiapatent.com.tw': 'TAIA',
    'wispro.com': 'Wispro',
    'adr-law.com': 'ADR',
    'rfrip.com': 'RFR',
    'smartbiggar.ca': 'SmartBiggar',
    'altitude-ip.com': 'AltitudeIP',
    'boehmert.de': 'Boehmert',
    'tsailee.com.tw': 'TsaiLee',
    'kspat.com': 'KASAN',
    'clarivate.com': 'Clarivate',
    'cpaglobal.com': 'CPA',
    'cpahkltd.com': 'CPAHK',
    'vjp.de': 'VJP',
    'naipo.com': 'NAIPO',
    'webmail.mozlen.com': 'Mozlen',
    'za3c.com.tw': 'ZA3C',
    'wglaw.com.tw': 'WGLaw',
    'taiwanlaw.com': 'TaiwanLaw',
}

IPWINNER_DOMAINS = ['ipwinner.com', 'ipwinner.com.tw']

# ============================================================
# === Email Parsing Functions ===
# ============================================================
def decode_header(raw):
    if raw is None:
        return ""
    decoded_parts = []
    for part, charset in email.header.decode_header(raw):
        if isinstance(part, bytes):
            try:
                decoded_parts.append(part.decode(charset or 'utf-8', errors='replace'))
            except (LookupError, UnicodeDecodeError):
                decoded_parts.append(part.decode('utf-8', errors='replace'))
        else:
            decoded_parts.append(str(part))
    return ' '.join(decoded_parts)

def extract_email_addr(raw):
    if not raw:
        return ""
    _, addr = email.utils.parseaddr(raw)
    return addr.lower()

def extract_all_recipients(msg):
    recipients = []
    for hdr in ['to', 'cc', 'bcc']:
        raw = msg.get(hdr, '')
        if raw:
            addrs = email.utils.getaddresses([raw])
            for _, addr in addrs:
                if addr:
                    recipients.append(addr.lower())
    return recipients

def get_body_text(msg, max_chars=500):
    body = ""
    if msg.is_multipart():
        for part in msg.walk():
            ct = part.get_content_type()
            if ct == 'text/plain':
                try:
                    payload = part.get_payload(decode=True)
                    charset = part.get_content_charset() or 'utf-8'
                    body = payload.decode(charset, errors='replace')
                    break
                except:
                    continue
            elif ct == 'text/html' and not body:
                try:
                    payload = part.get_payload(decode=True)
                    charset = part.get_content_charset() or 'utf-8'
                    html = payload.decode(charset, errors='replace')
                    body = re.sub(r'<[^>]+>', ' ', html)
                    body = re.sub(r'\s+', ' ', body).strip()
                except:
                    continue
    else:
        try:
            payload = msg.get_payload(decode=True)
            charset = msg.get_content_charset() or 'utf-8'
            if payload:
                body = payload.decode(charset, errors='replace')
        except:
            pass
    return body[:max_chars] if body else ""

def get_attachment_names(msg):
    attachments = []
    if msg.is_multipart():
        for part in msg.walk():
            if part.get_content_maintype() == 'multipart':
                continue
            filename = part.get_filename()
            if filename:
                attachments.append(decode_header(filename))
    return attachments

def clean_subject(subject):
    cleaned = re.sub(r'^(RE:|Re:|re:|Fwd:|FW:|fw:|回覆:|轉寄:|答覆:)\s*', '', subject)
    while cleaned != subject:
        subject = cleaned
        cleaned = re.sub(r'^(RE:|Re:|re:|Fwd:|FW:|fw:|回覆:|轉寄:|答覆:)\s*', '', subject)
    return cleaned.strip()

# Pattern for [From: "Name" <email>] or [From: Name <email>] prefix in subject
FROM_PREFIX_RE = re.compile(
    r'^\[From:\s*(?:"?([^"<]*)"?\s*)?<?([^>@\]]+@[^>\]]+)>?\]\s*',
    re.IGNORECASE
)

def extract_from_prefix(subject):
    """Extract embedded [From: ...] sender info from forwarded email subjects.
    Returns (real_sender_email, cleaned_subject) or (None, original_subject)."""
    m = FROM_PREFIX_RE.match(subject)
    if m:
        sender_email = m.group(2).strip().lower()
        real_subject = subject[m.end():].strip()
        # Recursively clean RE:/Fwd: from the remaining subject
        real_subject = clean_subject(real_subject)
        return sender_email, real_subject
    return None, subject

def extract_case_numbers(subject, body_snippet, attachment_names):
    all_text = f"{subject} {body_snippet} {' '.join(attachment_names)}"
    cases = list(set(CASE_RE.findall(all_text)))
    if not cases:
        base_cases = list(set(CASE_BASE_RE.findall(all_text)))
        cases = base_cases
    return cases

def extract_date_from_email(msg):
    """Extract date from email headers, return yyyymmdd string."""
    date_str = msg.get('date', '')
    if date_str:
        try:
            parsed = email.utils.parsedate_to_datetime(date_str)
            return parsed.strftime('%Y%m%d')
        except:
            pass
    # Fallback: try to extract from other headers or filename
    return '00000000'

# ============================================================
# === Classification Functions ===
# ============================================================
def get_sender_role(sender_domain):
    if sender_domain in ROLE_MAP_DOMAIN:
        return ROLE_MAP_DOMAIN[sender_domain]
    for known_domain, role in ROLE_MAP_DOMAIN.items():
        if sender_domain.endswith(known_domain):
            return role
    return 'X'

def get_recipient_role(recipients):
    for r in recipients:
        domain = r.split('@')[-1] if '@' in r else ''
        for known_domain, role in ROLE_MAP_DOMAIN.items():
            if domain == known_domain or domain.endswith(known_domain):
                if role != 'SELF':
                    return role
    return 'X'

def determine_send_receive_code(direction, sender_domain, recipients):
    if direction == 'F':
        role = get_sender_role(sender_domain)
        if role == 'C': return 'FC'
        if role == 'A': return 'FA'
        if role == 'G': return 'FG'
        return 'FX'
    else:
        role = get_recipient_role(recipients)
        if role == 'C': return 'TC'
        if role == 'A': return 'TA'
        if role == 'G': return 'TG'
        return 'TX'

def determine_case_category(case_numbers):
    has_patent = False
    has_trademark = False
    for cn in case_numbers:
        m = CASE_RE.match(cn)
        if m:
            full = m.group(1)
            base_match = re.match(r'[A-Z0-9]{4}\d{5}', full)
            if base_match:
                type_code = full[len(base_match.group(0)):len(base_match.group(0))+1]
                if type_code in 'PMD':
                    has_patent = True
                elif type_code == 'T':
                    has_trademark = True
    if has_patent and not has_trademark:
        return '專利'
    elif has_trademark and not has_patent:
        return '商標'
    elif has_patent and has_trademark:
        return '專利'
    return '未分類'

def generate_semantic_name(subject, direction, code, sender_domain, attachment_names, body_snippet):
    """Generate semantic filename using rule-based pattern matching."""
    subj = subject.strip()
    subj_lower = subj.lower()
    agent_code = AGENT_CODE_MAP.get(sender_domain, '')

    # === High-priority patterns ===

    # E-Filing Receipt / TIPO
    if 'e-filing receipt' in subj_lower or 'e-filing' in subj_lower:
        return 'E-Filing-Receipt'
    if 'tipo' in subj_lower and ('電子收據' in subj or 'receipt' in subj_lower):
        return 'TIPO電子收據'
    if '線上變更' in subj and '成功通知' in subj:
        return '線上變更成功通知'
    if 'tipo' in subj_lower and '帳單' in subj:
        return 'TIPO電子帳單'

    # Billing / 帳單
    if agent_code and ('invoice' in subj_lower or 'debit note' in subj_lower or '帳單' in subj or 'statement' in subj_lower):
        # Try to extract case type from subject for bracket code
        bracket = extract_bracket_code(subj, body_snippet, attachment_names)
        if bracket:
            return f'{agent_code}帳單-({bracket})'
        return f'{agent_code}帳單'

    # Filing report / 送件報告
    if 'filing report' in subj_lower or '送件報告' in subj:
        bracket = extract_bracket_code(subj, body_snippet, attachment_names)
        if bracket:
            return f'送件報告-({bracket})'
        return '送件報告'

    # OA related
    oa_match = re.search(r'(OA|office action|ROA)\s*[-#]?\s*(\d+)?', subj, re.I)
    if oa_match:
        oa_type = oa_match.group(1).upper()
        if oa_type == 'OFFICE ACTION':
            oa_type = 'OA'
        oa_num = oa_match.group(2) or '1'

        if any(kw in subj for kw in ['答辯', '指示']) or 'instruction' in subj_lower:
            if code.startswith('F') and code[1] == 'C':
                return f'答辯指示-({oa_type}{oa_num})'
            return f'委託{oa_type}{oa_num}答辯'
        if 'response' in subj_lower:
            if code.startswith('T'):
                return f'委託{oa_type}{oa_num}答辯'
            return f'回覆-({oa_type}{oa_num})'
        if '分析' in subj or 'analysis' in subj_lower:
            return f'{oa_type}{oa_num}分析'
        if '轉寄' in subj or 'forward' in subj_lower or subj_lower.startswith('fw:') or subj_lower.startswith('fwd:'):
            return f'轉寄-({oa_type}{oa_num})'
        if 'filing report' in subj_lower:
            return f'送件報告-({oa_type}{oa_num})'
        return f'{oa_type}{oa_num}相關'

    # 進度通知
    if '進度通知' in subj:
        m = re.search(r'進度通知[-\s]*(.+)', subj)
        if m:
            detail = m.group(1).strip()[:20]
            return f'進度通知-{detail}'
        return '進度通知'

    # Forward / 轉寄
    if subj_lower.startswith('fwd:') or subj_lower.startswith('fw:') or '轉寄' in subj:
        # Try to identify what's being forwarded
        inner = re.sub(r'^(Fwd:|FW:|fw:|轉寄[:：]?)\s*', '', subj, flags=re.I).strip()
        if inner:
            inner_clean = inner[:20]
            return f'轉寄-({inner_clean})'
        return '轉寄-官方文件'

    # 委託
    if '委託' in subj:
        bracket = extract_bracket_code(subj, body_snippet, attachment_names)
        m = re.search(r'委託(.+)', subj)
        if bracket:
            detail = m.group(1).strip()[:15] if m else ''
            return f'委託{detail}-({bracket})' if detail else f'委託-({bracket})'
        if m:
            return f'委託{m.group(1).strip()[:20]}'
        return '委託'

    # 確認 patterns
    if '確認承辦' in subj:
        return '確認承辦'
    if '確認可送件' in subj:
        return '確認可送件'
    if '確認' in subj:
        m = re.search(r'確認(.+)', subj)
        if m:
            return f'確認{m.group(1).strip()[:20]}'
        return '確認'

    # 領證
    if '領證' in subj:
        if '指示' in subj:
            return '領證指示'
        if '通知' in subj:
            return '專利領證通知'
        return '領證相關'

    # 年費
    if '年費' in subj or 'annuit' in subj_lower or 'maintenance fee' in subj_lower:
        if '指示' in subj:
            return '年費指示'
        if '通知' in subj or 'due' in subj_lower:
            return '年費到期通知'
        return '年費相關'

    # 商標 specific
    if '商標' in subj:
        if '監控' in subj:
            return '商標監控定期報告'
        if '核准' in subj:
            return '商標核准通知'
        if '註冊證' in subj:
            return '轉寄-(商標註冊證)'
        if '延展' in subj or '續展' in subj:
            return '商標延展'

    # Trademark watch/alert
    if 'trademark watch' in subj_lower or ('alert' in subj_lower and 'trademark' in subj_lower):
        return '商標監控通知'
    if 'markify' in subj_lower or (sender_domain == 'markify.com'):
        return '商標監控通知'

    # 已簽 / signed
    if '已簽' in subj or ('signed' in subj_lower and 'document' in subj_lower):
        return '已簽文件'

    # 提供
    if '提供' in subj:
        m = re.search(r'提供(.+)', subj)
        if m:
            return f'提供{m.group(1).strip()[:20]}'
        return '提供文件'

    # 詢問
    if '詢問' in subj:
        m = re.search(r'詢問(.+)', subj)
        if m:
            return f'詢問{m.group(1).strip()[:20]}'
        return '詢問'

    # 提醒
    if '提醒' in subj:
        return '提醒回覆指示'

    # 送核/校稿/draft
    if '送核' in subj:
        return '送核'
    if '校稿' in subj:
        return '校稿意見'
    if 'draft' in subj_lower and 'response' not in subj_lower:
        return '送核'

    # Registration/certificate
    if 'registration' in subj_lower or 'certificate' in subj_lower:
        if 'extend' in subj_lower or 'renew' in subj_lower:
            return '延展通知'
        return '註冊證書通知'

    # Allowance / grant
    if 'allowance' in subj_lower or 'grant' in subj_lower:
        return '核准通知'

    # Payment / 付款
    if '付款' in subj or 'payment' in subj_lower:
        return '付款相關'

    # Termination
    if 'terminat' in subj_lower and 'represent' in subj_lower:
        return '解除代理通知'

    # Power of Attorney
    if 'power of attorney' in subj_lower or 'POA' in subj:
        return 'POA相關'

    # IDS
    if 'ids' in subj_lower and ('submission' in subj_lower or 'filing' in subj_lower or '送件' in subj):
        return 'IDS送件'

    # Priority document
    if '優先權' in subj or 'priority' in subj_lower:
        return '優先權文件'

    # 說明 / explanation
    if '說明' in subj:
        m = re.search(r'說明(.+)', subj)
        if m:
            return f'說明{m.group(1).strip()[:20]}'
        return '說明'

    # 建議 / suggestion
    if '建議' in subj:
        return '建議'

    # 告知 / inform
    if '告知' in subj:
        m = re.search(r'告知(.+)', subj)
        if m:
            return f'告知{m.group(1).strip()[:20]}'
        return '告知'

    # 結案 / close case
    if '結案' in subj:
        return '結案通知'

    # 回覆
    if '回覆' in subj:
        m = re.search(r'回覆(.+)', subj)
        if m:
            return f'回覆{m.group(1).strip()[:20]}'
        return '回覆'

    # General fallback: use cleaned subject truncated
    if len(subj) > 0:
        # Remove problematic filesystem characters
        safe = re.sub(r'[/\\:*?"<>|]', '-', subj)
        safe = re.sub(r'\s+', ' ', safe).strip()
        return safe[:30]

    return '未命名信件'

def extract_bracket_code(subject, body_snippet, attachment_names):
    """Extract bracket event code from subject/body/attachments."""
    all_text = f"{subject} {' '.join(attachment_names)}"

    # Patent type codes
    if re.search(r'(新案|new\s*application|new\s*filing)', all_text, re.I):
        if re.search(r'(design|設計|新式樣)', all_text, re.I):
            return 'D-新'
        return 'P-新'

    # OA/ROA with number
    oa_m = re.search(r'(ROA|OA)\s*[-#]?\s*(\d+)', all_text, re.I)
    if oa_m:
        return f'{oa_m.group(1).upper()}{oa_m.group(2)}'

    # 領證
    if '領證' in all_text:
        return '領證'

    # 延展
    if '延展' in all_text or 'renew' in all_text.lower() or 'extend' in all_text.lower():
        return '延展'

    # Year fees
    yr_m = re.search(r'Y(\d+)', all_text)
    if yr_m:
        return f'Y{yr_m.group(1)}'

    # Translation
    if '翻譯' in all_text or 'translat' in all_text.lower():
        return '翻譯'

    # POA
    if 'POA' in all_text or 'power of attorney' in all_text.lower() or '委任狀' in all_text:
        return 'POA'

    # IDS
    if 'IDS' in all_text:
        return 'IDS'

    # 補正
    if '補正' in all_text or 'amendment' in all_text.lower():
        return '補正'

    # 商標新案
    if '商標' in all_text and ('新案' in all_text or 'new' in all_text.lower()):
        return 'T-新'

    return None

def sanitize_filename(name):
    """Make filename safe for filesystem."""
    # Remove/replace characters that are problematic
    safe = re.sub(r'[/\\:*?"<>|\x00-\x1f]', '-', name)
    safe = re.sub(r'-+', '-', safe)
    safe = safe.strip('- ')
    return safe[:80]  # Limit length

# ============================================================
# === Main Processing Pipeline ===
# ============================================================
def process_single_eml(filepath):
    """Parse and classify a single .eml file."""
    try:
        with open(filepath, 'rb') as f:
            msg = email.message_from_binary_file(f, policy=policy.compat32)

        raw_subject = decode_header(msg.get('subject', ''))
        sender_email = extract_email_addr(msg.get('from', ''))
        sender_domain = sender_email.split('@')[-1] if '@' in sender_email else ''
        recipients = extract_all_recipients(msg)
        cleaned_subject = clean_subject(raw_subject)
        body_snippet = get_body_text(msg, max_chars=500)
        attachment_names = get_attachment_names(msg)
        date_str = extract_date_from_email(msg)

        # Direction based on email headers
        direction = 'T' if sender_domain in IPWINNER_DOMAINS else 'F'

        # Handle [From: ...] prefix in forwarded emails
        # If direction=T and subject has [From: email@domain], the actual sender is embedded
        effective_sender_email = sender_email
        effective_sender_domain = sender_domain
        effective_direction = direction
        from_prefix_email, effective_subject = extract_from_prefix(cleaned_subject)

        if from_prefix_email:
            # This is a forwarded email; use the embedded sender for classification
            effective_sender_email = from_prefix_email
            effective_sender_domain = from_prefix_email.split('@')[-1] if '@' in from_prefix_email else ''
            # Re-determine direction based on original sender
            if effective_sender_domain in IPWINNER_DOMAINS:
                effective_direction = 'T'
            else:
                effective_direction = 'F'
            # Also re-clean the subject
            cleaned_subject = effective_subject if effective_subject else cleaned_subject
            # Further strip [From: ] from any remaining prefixed content
            while True:
                more_from, more_subj = extract_from_prefix(cleaned_subject)
                if more_from:
                    # If original sender was external, keep the first one
                    cleaned_subject = more_subj
                else:
                    break
            # Also strip 【請回覆】【供歸檔】 etc.
            cleaned_subject = re.sub(r'^【[^】]*】\s*', '', cleaned_subject).strip()

        # Case numbers
        case_numbers = extract_case_numbers(cleaned_subject, body_snippet, attachment_names)

        # Classification using effective sender info
        code = determine_send_receive_code(effective_direction, effective_sender_domain, recipients)
        semantic_name = generate_semantic_name(
            cleaned_subject, effective_direction, code, effective_sender_domain, attachment_names, body_snippet
        )
        category = determine_case_category(case_numbers)

        # Build case reference for filename
        if case_numbers:
            case_ref = case_numbers[0]  # Use first case number
            if len(case_numbers) > 1:
                case_ref += '等'
        else:
            case_ref = 'NOCASE'

        # Build new filename
        safe_semantic = sanitize_filename(semantic_name)
        new_filename = f"{date_str}-{code}-{case_ref}-{safe_semantic}.eml"

        return {
            'original_path': str(filepath),
            'original_filename': os.path.basename(filepath),
            'new_filename': new_filename,
            'date': date_str,
            'send_receive_code': code,
            'direction': effective_direction,
            'header_sender': sender_email,
            'effective_sender': effective_sender_email,
            'effective_sender_domain': effective_sender_domain,
            'had_from_prefix': from_prefix_email is not None,
            'recipients': recipients[:5],
            'raw_subject': raw_subject,
            'cleaned_subject': cleaned_subject,
            'case_numbers': case_numbers,
            'case_category': category,
            'semantic_name': semantic_name,
            'has_attachments': len(attachment_names) > 0,
            'attachment_count': len(attachment_names),
            'body_snippet_length': len(body_snippet),
            'error': None,
        }
    except Exception as e:
        return {
            'original_path': str(filepath),
            'original_filename': os.path.basename(filepath),
            'new_filename': None,
            'error': str(e),
        }

def main():
    print(f"=" * 60)
    print(f"V3 Email Processor - 下載區批量處理")
    print(f"=" * 60)
    print(f"輸入目錄: {INPUT_DIR}")
    print(f"輸出目錄: {OUTPUT_DIR}")
    print()

    # Collect all .eml files
    eml_files = []
    for root, dirs, files in os.walk(INPUT_DIR):
        for fname in sorted(files):
            if fname.lower().endswith('.eml'):
                eml_files.append(os.path.join(root, fname))

    print(f"找到 {len(eml_files)} 封 .eml 檔案")

    # Create output directory
    os.makedirs(OUTPUT_DIR, exist_ok=True)

    # Process each file
    results = []
    code_counter = Counter()
    category_counter = Counter()
    error_count = 0
    used_filenames = set()

    for i, fpath in enumerate(eml_files):
        if (i + 1) % 50 == 0 or i == 0:
            print(f"處理中: {i+1}/{len(eml_files)}...")

        result = process_single_eml(fpath)
        results.append(result)

        if result['error']:
            error_count += 1
            continue

        code_counter[result['send_receive_code']] += 1
        category_counter[result['case_category']] += 1

        # Handle filename collision with proper dedup
        new_name = result['new_filename']
        if new_name in used_filenames:
            base, ext = os.path.splitext(new_name)
            counter = 1
            while f"{base}_{counter}{ext}" in used_filenames:
                counter += 1
            new_name = f"{base}_{counter}{ext}"
            result['new_filename'] = new_name
        used_filenames.add(new_name)

        # Copy file with new name
        dest_path = os.path.join(OUTPUT_DIR, new_name)
        try:
            shutil.copy2(fpath, dest_path)
        except Exception as e:
            result['error'] = f'copy failed: {e}'
            error_count += 1

    # Print summary
    print()
    print(f"=" * 60)
    print(f"處理完成！")
    print(f"=" * 60)
    print(f"總數: {len(eml_files)}")
    print(f"成功: {len(eml_files) - error_count}")
    print(f"錯誤: {error_count}")
    print()

    print(f"--- 收發碼分布 ---")
    for code, count in sorted(code_counter.items()):
        pct = count / len(eml_files) * 100
        print(f"  {code}: {count:4d} ({pct:.1f}%)")

    print()
    print(f"--- 案件類別分布 ---")
    for cat, count in sorted(category_counter.items()):
        pct = count / len(eml_files) * 100
        print(f"  {cat}: {count:4d} ({pct:.1f}%)")

    # Save log
    log = {
        'summary': {
            'total': len(eml_files),
            'success': len(eml_files) - error_count,
            'errors': error_count,
            'code_distribution': dict(code_counter),
            'category_distribution': dict(category_counter),
            'output_dir': OUTPUT_DIR,
        },
        'results': results,
        'errors': [r for r in results if r['error']],
    }

    with open(LOG_PATH, 'w', encoding='utf-8') as f:
        json.dump(log, f, ensure_ascii=False, indent=2)

    print(f"\n分類日誌已儲存到: {LOG_PATH}")
    print(f"重新命名的 .eml 檔案在: {OUTPUT_DIR}")

    # Count output files
    out_files = [f for f in os.listdir(OUTPUT_DIR) if f.endswith('.eml')]
    print(f"輸出目錄中的 .eml 檔案數: {len(out_files)}")

if __name__ == '__main__':
    main()
