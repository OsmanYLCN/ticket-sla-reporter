#!/usr/bin/env python
# -*- coding: utf-8 -*-
"""
IT Infrastructure SLA Summary & Quarterly Consolidation Tool
============================================================
Consolidates quarterly ticket export sheets (Q1-Q4) into a 'Year-2026' sheet,
applies strict SLA and Scope categorization rules, and generates an executive
Summary dashboard with dynamic Excel formulas.

Usage:
    python generate_sla_summary.py
    python generate_sla_summary.py sample_data/tickets_export_20260831_084937.xlsx
    python generate_sla_summary.py -i sample_data/tickets_export_20260831_084937.xlsx -o outputs/sla_report.xlsx
"""

import sys
import os
import argparse
import time
from datetime import datetime

# Add package root to sys.path
sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))

from src.processor import load_and_consolidate_tickets
from src.excel_builder import build_sla_report_workbook

# Ensure UTF-8 console output on Windows
if sys.stdout.encoding != 'utf-8':
    try:
        sys.stdout.reconfigure(encoding='utf-8')
    except Exception:
        pass

# Enable ANSI Virtual Terminal Processing on Windows 10/11
if sys.platform == 'win32':
    try:
        import ctypes
        kernel32 = ctypes.windll.kernel32
        hOut = kernel32.GetStdHandle(-11)
        mode = ctypes.c_ulong()
        kernel32.GetConsoleMode(hOut, ctypes.byref(mode))
        kernel32.SetConsoleMode(hOut, mode.value | 0x0004)
    except Exception:
        pass


import re

ANSI_RE = re.compile(r'\033\[[0-9;]*m')


class UI:
    """Terminal styling, ANSI colors, and UI formatting helper."""
    IS_TTY = sys.stdout.isatty()
    
    RESET   = '\033[0m'  if IS_TTY else ''
    BOLD    = '\033[1m'  if IS_TTY else ''
    DIM     = '\033[2m'  if IS_TTY else ''
    
    BLUE    = '\033[94m' if IS_TTY else ''
    CYAN    = '\033[96m' if IS_TTY else ''
    GREEN   = '\033[92m' if IS_TTY else ''
    YELLOW  = '\033[93m' if IS_TTY else ''
    RED     = '\033[91m' if IS_TTY else ''
    WHITE   = '\033[97m' if IS_TTY else ''
    GRAY    = '\033[90m' if IS_TTY else ''

    @staticmethod
    def vlen(s):
        """Calculates visible length ignoring ANSI color escape sequences."""
        return len(ANSI_RE.sub('', str(s)))

    @staticmethod
    def pad(s, width, align='left'):
        """Pads string to exact visible width respecting ANSI escape sequences."""
        diff = max(0, width - UI.vlen(s))
        if align == 'right':
            return (' ' * diff) + s
        elif align == 'center':
            left = diff // 2
            right = diff - left
            return (' ' * left) + s + (' ' * right)
        return s + (' ' * diff)


def print_banner():
    """Prints a corporate header banner."""
    title = "IT INFRASTRUCTURE - SLA REPORTING & CONSOLIDATION ENGINE"
    print("\n" + f"{UI.CYAN}{UI.BOLD}╔" + "═" * 75 + "╗")
    print(f"║  {title:<73}║")
    print("╚" + "═" * 75 + f"╝{UI.RESET}\n")


def get_available_input_files():
    """Scans ONLY sample_data/ for all available raw .xlsx files sorted newest to oldest."""
    candidates = []
    sample_dir = os.path.join(os.path.dirname(os.path.abspath(__file__)), 'sample_data')
    
    if os.path.exists(sample_dir):
        for f in os.listdir(sample_dir):
            if f.endswith('.xlsx') and not f.startswith('~$'):
                full_p = os.path.join(sample_dir, f)
                rel_p = os.path.join('sample_data', f)
                mtime = os.path.getmtime(full_p)
                size_kb = os.path.getsize(full_p) / 1024.0
                candidates.append((f, full_p, rel_p, mtime, size_kb))
                
    # Sort descending by modification time (newest first)
    candidates.sort(key=lambda x: x[3], reverse=True)
    return candidates


def select_input_file_interactive():
    """Presents an interactive numbered list of raw Excel files inside sample_data/."""
    files = get_available_input_files()
    
    if not files:
        print(f"{UI.YELLOW}[!] 'sample_data/' klasöründe işlenecek Excel (.xlsx) dosyası bulunamadı.{UI.RESET}")
        print("    Lütfen ham Excel dosyanızı 'sample_data/' klasörüne ekleyip tekrar deneyin.")
        print(f"    {UI.RED}(Q){UI.RESET} Programdan çıkış")
        try:
            manual = input("\n    İşlenecek dosya yolunu elle giriniz [veya 'Q' ile çıkış]: ").strip()
        except (EOFError, KeyboardInterrupt):
            print(f"\n{UI.DIM}[*] Programdan çıkış yapıldı.{UI.RESET}\n")
            sys.exit(0)
            
        if manual.upper() == 'Q':
            print(f"\n{UI.DIM}[*] Programdan çıkış yapıldı.{UI.RESET}\n")
            sys.exit(0)
        return manual
        
    print(f"{UI.BOLD}[?] 'sample_data/' Klasöründeki Excel Dosyaları (Güncelden Eskiye):{UI.RESET}")
    print(f"{UI.DIM}  " + "─" * 73 + f"{UI.RESET}")
    
    for idx, (filename, full_p, rel_p, mtime, size_kb) in enumerate(files, start=1):
        dt_str = datetime.fromtimestamp(mtime).strftime('%Y-%m-%d %H:%M')
        size_str = f"{size_kb:,.0f} KB" if size_kb < 1024 else f"{size_kb/1024:,.1f} MB"
        tag = f" {UI.GREEN}[EN GÜNCEL]{UI.RESET}" if idx == 1 else ""
        print(f"  {UI.CYAN}({idx}){UI.RESET} {filename:<38} {UI.DIM}[{dt_str} | {size_str:>8}]{UI.RESET}{tag}")
        
    print(f"{UI.DIM}  " + "─" * 73 + f"{UI.RESET}")
    print(f"  {UI.YELLOW}(M){UI.RESET} Manuel dosya yolu gir...")
    print(f"  {UI.RED}(Q){UI.RESET} Çıkış (Quit)")
    
    while True:
        try:
            choice = input(f"\nLütfen işlenecek dosyayı seçin [1-{len(files)}, M, Q veya Enter: (1)]: ").strip()
        except (EOFError, KeyboardInterrupt):
            print(f"\n{UI.DIM}[*] Programdan güvenli şekilde çıkış yapıldı.{UI.RESET}\n")
            sys.exit(0)
            
        if choice.upper() == "Q":
            print(f"\n{UI.DIM}[*] Programdan güvenli şekilde çıkış yapıldı.{UI.RESET}\n")
            sys.exit(0)
        if choice == "" or choice == "1":
            selected = files[0][2]
            print(f"  -> {UI.GREEN}Seçilen:{UI.RESET} ({1}) {files[0][0]}")
            return selected
        if choice.upper() == "M":
            try:
                manual = input("  Lütfen Excel dosya yolunu giriniz [veya 'Q' ile iptal]: ").strip()
            except (EOFError, KeyboardInterrupt):
                print(f"\n{UI.DIM}[*] Programdan güvenli şekilde çıkış yapıldı.{UI.RESET}\n")
                sys.exit(0)
            if manual.upper() == "Q":
                print(f"\n{UI.DIM}[*] Programdan güvenli şekilde çıkış yapıldı.{UI.RESET}\n")
                sys.exit(0)
            if manual:
                return manual
        if choice.isdigit():
            val = int(choice)
            if 1 <= val <= len(files):
                selected = files[val - 1][2]
                print(f"  -> {UI.GREEN}Seçilen:{UI.RESET} ({val}) {files[val - 1][0]}")
                return selected
            else:
                print(f"  {UI.RED}[!] Geçersiz numara. Lütfen 1 ile {len(files)} arasında bir değer girin.{UI.RESET}")
        else:
            if os.path.exists(choice):
                return choice
            print(f"  {UI.RED}[!] Lütfen listeden bir numara (1-{len(files)}), 'M' veya 'Q' girin.{UI.RESET}")


def get_default_output_path(input_path):
    """Generates a clean destination path in outputs/ without touching original input."""
    base_name = os.path.splitext(os.path.basename(input_path))[0]
    out_dir = os.path.join(os.path.dirname(os.path.abspath(__file__)), 'outputs')
    os.makedirs(out_dir, exist_ok=True)
    return os.path.join('outputs', f"{base_name}_SLA_Summary.xlsx")


def select_output_file_interactive(input_path):
    """Prompts for output path with a smart default."""
    default_out = get_default_output_path(input_path)
    print(f"\n{UI.BOLD}[?] Çıktı Rapor Dosyası Hedefi:{UI.RESET}")
    print(f"    Varsayılan: {UI.CYAN}{default_out}{UI.RESET}")
    try:
        user_out = input(f"    Farklı bir çıktı yolu girmek için yazın, onaylamak için [Enter], çıkmak için [Q]: ").strip()
    except (EOFError, KeyboardInterrupt):
        return default_out
        
    if user_out.upper() == "Q":
        print(f"\n{UI.DIM}[*] Programdan güvenli şekilde çıkış yapıldı.{UI.RESET}\n")
        sys.exit(0)
    return user_out if user_out else default_out


def print_quarterly_panel(quarterly_counts, total_count):
    """Renders an executive ASCII status panel showing quarterly data counts."""
    W = 67
    print(f"\n{UI.BOLD}  ┌" + "─" * W + "┐")
    title = "ÇEYREK BİLET VERİLERİ (QUARTERLY DATA INGESTION)"
    print(f"  │  {UI.pad(title, W - 2)}│")
    print(f"  ├" + "─" * W + f"┤{UI.RESET}")
    
    for q_name in ['Tickets-Q1', 'Tickets-Q2', 'Tickets-Q3', 'Tickets-Q4']:
        cnt = quarterly_counts.get(q_name, 0)
        if cnt > 0:
            status_tag = f"{UI.GREEN}[✓] {q_name:<11} : {cnt:>6,} bilet (Yüklendi){UI.RESET}"
        else:
            status_tag = f"{UI.DIM}[ ] {q_name:<11} :      0 bilet (Boş / Veri Yok){UI.RESET}"
        print(f"  │  {UI.pad(status_tag, W - 2)}│")
        
    print(f"{UI.BOLD}  ├" + "─" * W + "┤")
    tot_label = f"Toplam Konsolide Edilen Bilet : {total_count:,}"
    print(f"  │  {UI.pad(f'{UI.CYAN}{tot_label}{UI.RESET}', W - 2)}{UI.BOLD}│")
    print(f"  └" + "─" * W + f"┘{UI.RESET}\n")


def print_completion_card(input_path, output_path, total_records, stats, elapsed):
    """Renders the final executive summary card with metrics and SLA results."""
    in_scope = stats.get('in_scope_count', 0)
    max_row = stats.get('max_row', total_records + 1)
    net_cnt = stats.get('net_count', 0)
    ot_cnt = stats.get('ot_count', 0)
    passed = stats.get('passed_count', 0)
    failed = stats.get('failed_count', 0)
    
    rate = (passed / in_scope * 100) if in_scope > 0 else 0.0
    sla_result = "PASSED" if rate >= 90.0 else "FAILED"
    sla_color = UI.GREEN if sla_result == "PASSED" else UI.RED
    
    W = 88
    header_text = "İŞLEM BAŞARIYLA TAMAMLANDI - YÖNETİCİ ÖZET RAPORU"
    header_padded = UI.pad(header_text, W, align='center')
    
    print("\n" + f"{UI.BOLD}╔" + "═" * W + "╗")
    print(f"║{UI.GREEN}{header_padded}{UI.RESET}{UI.BOLD}║")
    print("╠" + "═" * W + "╣")
    
    def row(label, val):
        padded_lbl = UI.pad(label, 25)
        content = f"  {padded_lbl} : {val}"
        print(f"║{UI.pad(content, W)}║")
        
    row(f"{UI.DIM}Giriş Dosyası (Ham){UI.RESET}", f"{input_path}")
    row(f"{UI.CYAN}Üretilen Rapor{UI.RESET}", f"{output_path}")
    print("╟" + "─" * W + "╢")
    row(f"{UI.BOLD}Toplam İncelenen Bilet{UI.RESET}", f"{total_records:>6,} bilet")
    row(f"{UI.YELLOW}Kapsama Giren Bilet{UI.RESET}", f"{in_scope:,} bilet  ({net_cnt} Network, {ot_cnt} Industrial OT)")
    row(f"{UI.BOLD}SLA Hedef Uyumu{UI.RESET}", f"{passed:,} Met (Passed)  /  {failed:,} Breached (Failed)")
    rate_str = f"{rate:.1f}%  (KPI Hedefi: %90.0)  [{sla_color}{UI.BOLD}{sla_result}{UI.RESET}{UI.BOLD}]"
    row(f"{UI.BOLD}SLA Başarı Oranı{UI.RESET}", rate_str)
    row(f"{UI.DIM}'Year-2026' Satır Sayısı{UI.RESET}", f"{max_row:>6,} satır")
    row(f"{UI.DIM}Toplam İşlem Süresi{UI.RESET}", f"{elapsed:.2f} saniye")
    print("╚" + "═" * W + f"╝{UI.RESET}\n")


def main():
    parser = argparse.ArgumentParser(
        description="IT Infrastructure - Ticket SLA Summary Reporter",
        formatter_class=argparse.RawTextHelpFormatter
    )
    parser.add_argument(
        'input_file',
        nargs='?',
        default=None,
        help="Path to the input Excel file (e.g., sample_data/tickets_export_20260831_084937.xlsx)"
    )
    parser.add_argument(
        '-i', '--input',
        dest='input_opt',
        default=None,
        help="Input Excel file path"
    )
    parser.add_argument(
        '-o', '--output',
        dest='output_file',
        default=None,
        help="Output Excel report path (defaults to outputs/<input_name>_SLA_Summary.xlsx)"
    )
    parser.add_argument(
        '--non-interactive',
        action='store_true',
        help="Run without interactive user prompts"
    )
    
    args = parser.parse_args()
    
    print_banner()
    
    # 1. Determine input path
    input_path = args.input_opt or args.input_file
    
    if not input_path:
        if not args.non_interactive and sys.stdin.isatty():
            input_path = select_input_file_interactive()
        else:
            files = get_available_input_files()
            if files:
                input_path = files[0][2]
            else:
                input_path = 'sample_data/tickets_export_20260831_084937.xlsx'
            print(f"[*] Giriş dosyası otomatik seçildi: {input_path}")
            
    if not os.path.exists(input_path):
        print(f"\n{UI.RED}[!] HATA: Belirtilen giriş dosyası bulunamadı: {input_path}{UI.RESET}\n")
        sys.exit(1)
        
    # 2. Determine output path
    output_path = args.output_file
    if not output_path:
        if not args.non_interactive and sys.stdin.isatty():
            output_path = select_output_file_interactive(input_path)
        else:
            output_path = get_default_output_path(input_path)
            
    # Ensure output directory exists
    out_dir = os.path.dirname(output_path)
    if out_dir:
        os.makedirs(out_dir, exist_ok=True)
        
    # Ensure input and output are distinct files to preserve the raw input
    if os.path.abspath(input_path) == os.path.abspath(output_path):
        print(f"\n{UI.YELLOW}[!] UYARI: Orijinal dosyanın bozulmaması için girdi ve çıktı aynı dosya olamaz.{UI.RESET}")
        output_path = get_default_output_path(input_path)
        print(f"    Çıktı güvenli konuma yönlendirildi: {output_path}")

    start_time = time.time()
    
    print(f"{UI.CYAN}[1/3]{UI.RESET} Orijinal dosya salt-okunur olarak taranıyor: {UI.BOLD}{input_path}{UI.RESET}")
    records, header_cols, quarterly_counts = load_and_consolidate_tickets(input_path)
    
    # Render Quarterly Ingestion Status Box
    print_quarterly_panel(quarterly_counts, len(records))
    
    if not records:
        print(f"{UI.RED}[!] HATA: Excel dosyasından hiçbir bilet verisi okunamadı.{UI.RESET}\n")
        sys.exit(1)
        
    print(f"{UI.CYAN}[2/3]{UI.RESET} Dinamik SLA formülleri ve 'Summary' Dashboard oluşturuluyor...")
    stats = build_sla_report_workbook(
        all_records=records,
        header_cols=header_cols,
        input_path=input_path,
        output_path=output_path,
        total_source_count=len(records)
    )
    
    elapsed = time.time() - start_time
    print(f"{UI.CYAN}[3/3]{UI.RESET} Çıktı Excel dosyası başarıyla kaydedildi: {UI.GREEN}{output_path}{UI.RESET}")
    
    # Render Executive Completion Card
    print_completion_card(input_path, output_path, len(records), stats, elapsed)

    # Prompt to open report in Microsoft Excel (Windows only)
    if sys.platform == 'win32' and not args.non_interactive and sys.stdin.isatty():
        try:
            open_resp = input(f"  {UI.BOLD}Rapor dosyasını Excel ile açmak ister misiniz?{UI.RESET} [E/H, Varsayılan: E]: ").strip().upper()
            if open_resp in ['', 'E', 'EVET', 'Y', 'YES']:
                print(f"  -> {UI.GREEN}Excel başlatılıyor...{UI.RESET}\n")
                os.startfile(os.path.abspath(output_path))
        except (EOFError, KeyboardInterrupt):
            pass


if __name__ == '__main__':
    try:
        main()
    except KeyboardInterrupt:
        print(f"\n\n{UI.DIM}[*] İşlem kullanıcı tarafından iptal edildi (Ctrl+C). Güvenli çıkış yapıldı.{UI.RESET}\n")
        sys.exit(0)
