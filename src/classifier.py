"""
Classification logic for Network, Industrial OT, and SLA scope rules.
"""

from .config import (
    TARGET_BUSINESS_LINE,
    TARGET_SEVERITIES,
    NET_KW_PATTERN,
    HQ_CLOUD_PATTERN,
    PLANT_PATTERN,
    OT_TECH_PATTERN,
    SAP_PATTERN,
    SAP_CONN_PATTERN
)


def get_plant_location(row_dict):
    """
    Determines the specific manufacturing plant/site for breakdown reporting.
    """
    loc_ci = str(row_dict.get('LocalitaCI', '') or '').lower()
    loc_user = str(row_dict.get('Ubicazione sede Utente', '') or '').lower()
    oggetto = str(row_dict.get('Oggetto', '') or '').lower()
    ci = str(row_dict.get('Configuration Item', '') or '').lower()
    desc = str(row_dict.get('Descrizione Articolo', '') or '').lower()
    
    full_loc = f"{loc_ci} {loc_user} {oggetto} {ci} {desc}"
    
    if any(k in full_loc for k in ['kocaeli', 'izmit', 'turkey', 'tr-tpi', 'dmctr']):
        return 'Kocaeli'
    if any(k in full_loc for k in ['alexandria', 'alessandria', 'egypt', 'eg-tpi', 'egals']):
        return 'Alexandria'
    if any(k in full_loc for k in ['santo andr', 'santo andre', 'brsa', 'fw2-sne']):
        return 'Santo André'
    if any(k in full_loc for k in ['gravata', 'brgv', 'srvbr045']):
        return 'Gravataí'
    return 'Other Sites'


def classify_ticket(row_dict):
    """
    Classifies a ticket according to Prometeon SLA rules.
    Returns: (in_scope: bool, scope_category: str)
    """
    bl = str(row_dict.get('BusinessLine', '') or '').strip()
    sev = str(row_dict.get('Severity', '') or '').strip()
    
    # Rule 1: Sadece BusinessLine == "Infrastructure Services"
    if bl != TARGET_BUSINESS_LINE:
        return False, 'Out of Scope - BusinessLine'
        
    # Rule 2: Sadece Severity == "1 - Molto alta" ve "2 - Alta"
    if sev not in TARGET_SEVERITIES:
        return False, 'Out of Scope - Severity'
        
    servizio = str(row_dict.get('Servizio', '') or '')
    tipologia = str(row_dict.get('Tipologia Calcolata', '') or '')
    oggetto = str(row_dict.get('Oggetto', '') or '')
    ci = str(row_dict.get('Configuration Item', '') or '')
    desc = str(row_dict.get('Descrizione Articolo', '') or '')
    loc_ci = str(row_dict.get('LocalitaCI', '') or '')
    loc_user = str(row_dict.get('Ubicazione sede Utente', '') or '')
    
    text_corpus = f"{oggetto} {servizio} {tipologia} {ci} {desc}"
    loc_corpus = f"{loc_ci} {loc_user}"
    
    # Rule 3: Network Check
    is_net_service_or_tipo = ('MANAGED NETWORK SERVICE' in servizio.upper()) or ('NETWORKING' in tipologia.upper())
    has_net_kw = bool(NET_KW_PATTERN.search(text_corpus))
    is_network = is_net_service_or_tipo or has_net_kw
    
    # Rule 5: SAP Exception Check
    has_sap = bool(SAP_PATTERN.search(text_corpus))
    is_sap_conn = has_sap and bool(SAP_CONN_PATTERN.search(text_corpus))
    is_sap_standard_error = has_sap and not is_sap_conn
    
    # Rule 4: Industrial OT Check
    # Exclude HQ (Cesano Maderno, Milano, TIM/HQ, Cinisello) and Cloud (Azure, GCP)
    is_hq_or_cloud = bool(HQ_CLOUD_PATTERN.search(loc_corpus)) or any(
        hq in loc_ci.lower() or hq in loc_user.lower()
        for hq in ['cesano', 'milano', 'cinisello', 'tim', 'hq', 'azure', 'gcp']
    )
    
    is_plant = bool(PLANT_PATTERN.search(loc_corpus)) or bool(PLANT_PATTERN.search(text_corpus))
    has_ot_tech = bool(OT_TECH_PATTERN.search(text_corpus))
    
    is_industrial_ot = (not is_hq_or_cloud) and is_plant and has_ot_tech and not is_sap_standard_error
    
    # Final Scope Decision
    if is_sap_standard_error:
        return False, 'Excluded - SAP Standard App'
    if is_network:
        return True, 'Network'
    if is_industrial_ot:
        return True, 'Industrial OT'
        
    return False, 'Out of Scope - Other Infra'
