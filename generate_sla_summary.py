#!/usr/bin/env python
# -*- coding: utf-8 -*-
"""
Prometeon IT Infrastructure SLA Summary & Quarterly Consolidation Tool
======================================================================
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
                candidates.append((f, full_p, rel_p, mtime))
                
    # Sort descending by modification time (newest first)
    candidates.sort(key=lambda x: x[3], reverse=True)
    return candidates


def select_input_file_interactive():
    """Presents an interactive numbered list of raw Excel files inside sample_data/."""
    files = get_available_input_files()
    
    if not files:
        print("\n[!] 'sample_data/' klasöründe işlenecek Excel (.xlsx) dosyası bulunamadı.")
        print("    Lütfen ham Excel dosyanızı 'sample_data/' klasörüne ekleyip tekrar deneyin.")
        print("    (Q) Programdan çıkış")
        manual = input("\n    İşlenecek dosya yolunu elle giriniz [veya 'Q' ile çıkış]: ").strip()
        if manual.upper() == 'Q':
            print("\n[*] Programdan çıkış yapıldı.\n")
            sys.exit(0)
        return manual
        
    print("\n[?] 'sample_data/' Klasöründeki Excel Dosyaları (Güncelden Eskiye Sıralı):")
    for idx, (filename, full_p, rel_p, mtime) in enumerate(files, start=1):
        dt_str = datetime.fromtimestamp(mtime).strftime('%Y-%m-%d %H:%M:%S')
        tag = " [EN GÜNCEL]" if idx == 1 else ""
        print(f"    ({idx}) {filename:<45} (Son Değişiklik: {dt_str}){tag}")
    print(f"    (M) Manuel dosya yolu girmek istiyorum")
    print(f"    (Q) Çıkış (Quit)")
    
    while True:
        try:
            choice = input(f"\nLütfen işlenecek dosyayı seçin [1-{len(files)}, M, Q veya Enter: (1)]: ").strip()
        except (EOFError, KeyboardInterrupt):
            print("\n[*] Programdan güvenli şekilde çıkış yapıldı.\n")
            sys.exit(0)
            
        if choice.upper() == "Q":
            print("\n[*] Programdan güvenli şekilde çıkış yapıldı.\n")
            sys.exit(0)
        if choice == "" or choice == "1":
            selected = files[0][2]
            print(f"    -> Seçilen: ({1}) {files[0][0]}")
            return selected
        if choice.upper() == "M":
            try:
                manual = input("    Lütfen Excel dosya yolunu giriniz [veya 'Q' ile iptal]: ").strip()
            except (EOFError, KeyboardInterrupt):
                print("\n[*] Programdan güvenli şekilde çıkış yapıldı.\n")
                sys.exit(0)
            if manual.upper() == "Q":
                print("\n[*] Programdan güvenli şekilde çıkış yapıldı.\n")
                sys.exit(0)
            if manual:
                return manual
        if choice.isdigit():
            val = int(choice)
            if 1 <= val <= len(files):
                selected = files[val - 1][2]
                print(f"    -> Seçilen: ({val}) {files[val - 1][0]}")
                return selected
            else:
                print(f"    [!] Geçersiz numara. Lütfen 1 ile {len(files)} arasında bir değer girin.")
        else:
            if os.path.exists(choice):
                return choice
            print(f"    [!] Lütfen listeden bir numara (1-{len(files)}), 'M' veya 'Q' girin.")


def get_default_output_path(input_path):
    """Generates a clean destination path in outputs/ without touching original input."""
    base_name = os.path.splitext(os.path.basename(input_path))[0]
    out_dir = os.path.join(os.path.dirname(os.path.abspath(__file__)), 'outputs')
    os.makedirs(out_dir, exist_ok=True)
    return os.path.join('outputs', f"{base_name}_SLA_Summary.xlsx")


def select_output_file_interactive(input_path):
    """Prompts for output path with a smart default."""
    default_out = get_default_output_path(input_path)
    print(f"\n[?] Çıktı Rapor Dosyası Konumu:")
    print(f"    Varsayılan Hedef: {default_out}")
    try:
        user_out = input(f"    Farklı bir çıktı yolu girmek için yazın, onaylamak için [Enter], çıkmak için [Q]: ").strip()
    except (EOFError, KeyboardInterrupt):
        return default_out
    if user_out.upper() == "Q":
        print("\n[*] Programdan güvenli şekilde çıkış yapıldı.\n")
        sys.exit(0)
    return user_out if user_out else default_out


def main():
    parser = argparse.ArgumentParser(
        description="Prometeon IT Infrastructure - Ticket SLA Summary Reporter",
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
    
    print("\n" + "="*75)
    print("  PROMETEON IT INFRASTRUCTURE - SLA SUMMARY & QUARTERLY REPORT GENERATOR")
    print("="*75)
    
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
            print(f"\n[*] Giriş dosyası otomatik seçildi: {input_path}")
            
    if not os.path.exists(input_path):
        print(f"\n[!] HATA: Belirtilen giriş dosyası bulunamadı: {input_path}")
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
        print(f"\n[!] UYARI: Orijinal dosyanın bozulmaması için girdi ve çıktı aynı dosya olamaz.")
        output_path = get_default_output_path(input_path)
        print(f"    Çıktı güvenli konuma yönlendirildi: {output_path}")

    start_time = time.time()
    print(f"\n[1/3] Orijinal dosya salt-okunur olarak taranıyor: {input_path}")
    records, header_cols, quarterly_counts = load_and_consolidate_tickets(input_path)
    
    for q_name, count in quarterly_counts.items():
        print(f"      - {q_name}: {count:,} bilet")
    print(f"      -> Toplam Konsolide Edilen Bilet Sayısı: {len(records):,}")
    
    if not records:
        print("[!] HATA: Excel dosyasından hiçbir bilet verisi okunamadı.")
        sys.exit(1)
        
    print(f"\n[2/3] Dinamik SLA formülleri ve 'Summary' Dashboard oluşturuluyor...")
    in_scope_count, max_row = build_sla_report_workbook(
        all_records=records,
        header_cols=header_cols,
        input_path=input_path,
        output_path=output_path,
        total_source_count=len(records)
    )
    
    elapsed = time.time() - start_time
    print(f"[3/3] Çıktı Excel dosyası başarıyla kaydedildi: {output_path}")
    
    print("\n" + "-"*75)
    print("  İŞLEM BAŞARIYLA TAMAMLANDI (ÖZET SONUÇLAR)")
    print("-"*75)
    print(f"  * Orijinal Giriş Dosyası (Korundu) : {input_path}")
    print(f"  * Üretilen Rapor Dosyası           : {output_path}")
    print(f"  * Toplam İncelenen Bilet Sayısı    : {len(records):,}")
    print(f"  * Kapsama Giren Bilet Sayısı (AH)  : {in_scope_count:,}")
    print(f"  * 'Year-2026' Satır Sayısı         : {max_row:,}")
    print(f"  * İşlem Süresi                     : {elapsed:.2f} saniye")
    print("="*75 + "\n")

    # Prompt to open report in Microsoft Excel (Windows only)
    if sys.platform == 'win32' and not args.non_interactive and sys.stdin.isatty():
        try:
            open_resp = input("  Rapor dosyasını Excel ile açmak ister misiniz? [E/H, Varsayılan: E]: ").strip().upper()
            if open_resp in ['', 'E', 'EVET', 'Y', 'YES']:
                print("  -> Excel başlatılıyor...\n")
                os.startfile(os.path.abspath(output_path))
        except Exception:
            pass


if __name__ == '__main__':
    try:
        main()
    except KeyboardInterrupt:
        print("\n\n[*] İşlem kullanıcı tarafından iptal edildi (Ctrl+C). Güvenli çıkış yapıldı.\n")
        sys.exit(0)
