import streamlit as st
import streamlit.components.v1 as components

APP_TITLE = "KP Sanghvi & Sons LLP"
APP_ICON = "https://kpsanghvi.com/wp-content/uploads/elementor/thumbs/kp-logo-l-1-ra4wu10y8udmdikjwyopdd9oovnzc0j2fiff236cvw.png"

SENTINEL = "🔍"
HOME = "Home"
NAV_VLOOKUP = "🔍 VLOOKUP"
NAV_MERGE = "🧾 Merge"
NAV_FEEDING = "📤 Feeding"

USERNAME = st.secrets.get("USERNAME")
PASSWORD = st.secrets.get("PASSWORD")

st.set_page_config(page_title=APP_TITLE, page_icon=APP_ICON, layout="wide")
from s_vinod import run_s_vinod_mod
from revision_stock import revision_stock_list
from print_sticker import run_print_sticker
from black_vlookpu import run_black_deal_mod
from all_black import run_black_mod
from browns_file import run_browns_mod
from web_fromat import run_web_fromat
from dela_disc_mod import run_dela_disc_mod
from vinay_diamond_black import run_vinay_black_mod
from gemstar_black import run_gemstar_black_mod
from unique_black import run_unique_black_mod
from diaclust_black import run_diaclust_black_mod
from all_deals import all_deal
from pdf_modify import pdf_modify_run
from sale_file import sale_file
from dossier_report import dossier_report
from rap_report import rap_report
from dossier_report_pdf import dosser_run
from rap_report_pdf import rap_report_pdf
from run_ratnakala_fancy import run_ratnakala_fancy_mod
from mpb_bill import run_mpb_bill
from hemil_check import run_hemil_check
from stock_sold_days import run_stock_list
from market_disc import run_market_disc
from revsion import run_revsion
from gia_price import run_gia_price
from pending_stock import run_pending_stock
from convert_stock_list import run_convert_stock_mod
from highlight_stone import run_highlight
from deal_Vlookup import run_discount_vlookup
from vinay_diamond import run_vinay_diamonds_mod
from veer_gems import run_veer_gems_mod
from vd_global import run_vd_global_mod
from vaibhav import run_vaibhav_mod
from unique import run_unique_mod
from sunrays import run_sunrays_mod
from starrays import run_starrays_mod
from starlight import run_starlight_mod
from srd import run_srd_mod
from snj import run_snj_mod
from ske import run_ske_mod
from shyam_copration import run_shyam_mod
from shivam_list import run_shivan_mod
from sheetal import run_sheetal_mod
from shairu import run_shairu_mod
from sanghvi_sons import run_sanghvi_sons_mod
from sagar_list import run_sagar_list_mod
from red_exim import run_red_exim_mod
from rays_web import run_rays_web_mod
from rays_mail import run_rays_mod
from ratnakala_web import run_ratnakala_web_mod
from ratnakala_mail import run_ratnakala_mail_mod
from r_rajesh import run_r_rajesh_mod
from r_n_export import run_rn_export_mod
from pramukh import run_pramukh_mod
from pragati import run_pragati_mod
from parishi import run_parishi_mod
from pansurys import run_pansurya_mod
from paladiya import run_paladiya_mod
from narola import run_narola_mod
from mb import run_mb_mod
from m_suresh import run_m_suresh_mod
from md_brother import run_md_brothers_mod
from akarsh import run_akarsh_mod
from amora import run_amora_mod
from ankit import run_ankit_mod
from asian import run_asian_mod
from d_navin import run_d_navin_mod
from diaclust import run_diaclust_mod
from feeding import run_feeding_mod  
from finestar import run_finestar_mod
from glowstar import run_glowstar_mod
from gyanimpex import run_gyanimpex_mod
from hrv import run_hrv_mod
from hk import run_hk_mod
from hiradias import run_hiradias_mod
from jewelex import run_jewelex_mod
from jodhani import run_jodhani_mod
from kb import run_kb_mod
from kbs_diamond import run_kbs_mod
from kbs_web import run_kbs_web_mod
from kc_mail import run_kc_mail_mod
from kc_web import run_kc_web_mod
from kg import run_kg_mod
from kgk import run_kgk_mod
from kiran import run_kiran_web_mod
from kiran_mail import run_kiran_mail_mod
from komal import run_komal_mod
from lakhi import run_lakhi_mod
from laxmi import run_laxmi_mod
from merge import run_merge_mod
from vlookup import run_vlookup_mod
from export import run_export

from ahs import run_ahs_mod
from mpb_nattes import run_mpb_natts_mod
from rapnet_format import run_rapnet_format_mod
from LED import run_led_mod
from canadian import run_canadian_mod
from veer_gems_black import run_veer_gems_black_mod
from vd_global_black import run_vd_global_black_mod
from vaibhav_black import run_vaibhav_black_mod
from sunrays_black import run_sunrays_black_mod
from starrays_black import run_starrays_black_mod
from starlight_black import run_starlight_black_mod
from srd_black import run_srd_black_mod
from snj_black import run_snj_black_mod
from shyam_black import run_shyam_black_mod
from shree_krishna_black import run_shree_krishna_black_mod
from shivam_black import run_shivam_black_mod
from sheetal_black import run_sheetal_black_mod
from shairu_black import run_shairu_black_mod
from Sanghvi_sons_black import run_sanghvi_sons_black_mod
from Sagar_black import run_sagar_black_mod
from red_exim_black import run_red_exim_black_mod
from rays_black import run_rays_black_mod
from ratnakala_black import run_ratnakala_black_mod
from r_rajesh_black import run_r_rajesh_black_mod
from r_n_black import run_r_n_black_mod
from pragati_black import run_pragati_black_mod
from pansurya_black import run_pansurya_black_mod
from paladiya_black import run_paladiya_black_mod
from narola_black import run_narola_black_mod
from pramukh_black import run_pramukh_black_mod
from m_suresh_black import run_m_suresh_black_mod
from mb_brother_black import run_mb_black_mod
from lakhi_black import run_lakhi_black_mod
from laxmi_black import run_laxmi_black_mod
from komal_black import run_komal_black_mod
from kiran_black import run_kiran_black_mod
from kgk_black import run_kgk_black_mod
from kg_black import run_kg_black_mod
from kc_black import run_kc_black_mod
from kbs_black import run_kbs_black_mod
from kb_black import run_kb_black_mod
from jodhani_black import run_jodhani_black_mod
from jewelex_black import run_jewelex_black_mod
from hrv_black import run_hrv_black_mod
from hk_black import run_hk_black_mod
from hiradias_black import run_hiradias_black_mod
from gnyanimpex_black import run_gnyan_black_mod
from glowstar_black import run_glowstar_black_mod
from finstar_black import run_finstar_black_mod
from d_navin_black import run_d_navin_black_mod
from asian_black import run_asian_black_mod
from amora_black import run_amora_black_mod
from ankit_black import run_ankit_black_mod
from akarsh_black import run_akarsh_black_mod
from kiran_fancy import run_kiran_fancy_mod
from lakhi_fancy import run_lakhi_fancy_mod
from harekrishna import run_harekrishna_mod
from rapnet import run_rapnet_mod
from ske import run_ske_mod
from srk import run_srk_mod

# -------------------------------
# REGISTRIES
# -------------------------------
REPORT_FUNCTION_MAP = {
    "Single Black File": run_black_deal_mod,
    "All Black": run_black_mod,
    "Akarsh Black": run_akarsh_black_mod,
    "Amora Black": run_amora_black_mod,
    "Ankit Black": run_ankit_black_mod,
    "Asian Black": run_asian_black_mod,
    "D Navin Black": run_d_navin_black_mod,
    "Diaclust Black": run_diaclust_black_mod,
    "Fine Star Black": run_finstar_black_mod,
    "Glow Satr Black": run_glowstar_black_mod,
    "Gem Star Black": run_gemstar_black_mod,
    "Gnyan Impex Black": run_gnyan_black_mod,
    "Hiradias Black": run_hiradias_black_mod,
    "HK Black": run_hk_black_mod,
    "HRV Black": run_hrv_black_mod,
    "Jewelex Black": run_jewelex_black_mod,
    "Jodhani Black": run_jodhani_black_mod,
    "KB Black": run_kb_black_mod,
    "KBS Black": run_kbs_black_mod,
    "KC Black": run_kc_black_mod,
    "KG Black": run_kg_black_mod,
    "KGK Black": run_kgk_black_mod,
    "Kiran Black": run_kiran_black_mod,
    "Komal Black": run_komal_black_mod,  
    "Laxmi Black": run_laxmi_black_mod,  
    "Lakhi Black": run_lakhi_black_mod,  
    "MB Black": run_mb_black_mod,  
    "M Suresh Black": run_m_suresh_black_mod,  
    "Narola Black": run_narola_black_mod,  
    "Pramukh Black": run_pramukh_black_mod,  
    "Paladiya Black": run_paladiya_black_mod,  
    "Pansurya Black": run_pansurya_black_mod,  
    "Pragati Black": run_pragati_black_mod,  
    "R N Export Black": run_r_n_black_mod,  
    "R Rajesh Black": run_r_rajesh_black_mod,  
    "Ratnakala Black": run_ratnakala_black_mod,
    "Rays Black": run_rays_black_mod,  
    "Red Exim Black": run_red_exim_black_mod,  
    "Sagar Black": run_sagar_black_mod,  
    "Sanghvi & Sons Black": run_sanghvi_sons_black_mod,  
    "Shairu Black": run_shairu_black_mod,  
    "Sheetal Black": run_sheetal_black_mod,  
    "Shivam Black": run_shivam_black_mod,  
    "Shree krishna Black": run_shree_krishna_black_mod,
    "Shyam Coperation Black": run_shyam_black_mod,  
    "SNJ Black": run_snj_black_mod,
    "SRD Black": run_srd_black_mod,
    "Starlight Black": run_starlight_black_mod,
    "Starrays Black": run_starrays_black_mod,
    "Sunrays Black": run_sunrays_black_mod,
    "Unique Black": run_unique_black_mod,
    "Vaibhav Black": run_vaibhav_black_mod,
    "VD Global Black": run_vd_global_black_mod,
    "Veer Gems Black": run_veer_gems_black_mod,
    "Vinay Diamond Black": run_vinay_black_mod,
}

DEAL_FUNCTION_MAP = {
    "Disc Vlookup":run_discount_vlookup,
    "All DEALS":all_deal,
    "Single DEALS File":run_dela_disc_mod,
    "Akarsh": run_akarsh_mod,
    "Amora": run_amora_mod,
    "Ankit": run_ankit_mod,
    "Asian Star": run_asian_mod,
    "D Navin": run_d_navin_mod,
    "Diaclust": run_diaclust_mod,
    "Finestar": run_finestar_mod,
    "Glowstar": run_glowstar_mod,
    "Gyanimpex": run_gyanimpex_mod,
    "Haridas": run_hiradias_mod,
    "Harekrishna":run_harekrishna_mod,
    "HK": run_hk_mod,
    "Hrv": run_hrv_mod,
    "Jewelex": run_jewelex_mod,
    "Jodhani": run_jodhani_mod,
    "KB": run_kb_mod,
    "KBS": run_kbs_mod,
    "KBS Web": run_kbs_web_mod,
    "KC Mail": run_kc_mail_mod,
    "KC Web": run_kc_web_mod,
    "KG": run_kg_mod,
    "KGK": run_kgk_mod,
    "Kiran Web": run_kiran_web_mod,
    "Kiran Mail": run_kiran_mail_mod,
    "Komal": run_komal_mod,
    "Lakhi": run_lakhi_mod,
    "Laxmi": run_laxmi_mod,
    "M D Brother": run_md_brothers_mod,
    "M Suresh": run_m_suresh_mod,
    "MB": run_mb_mod,
    "Narola": run_narola_mod,
    "Paladiya": run_paladiya_mod,
    "Pansurya": run_pansurya_mod,
    "parishi": run_parishi_mod,
    "Pragati": run_pragati_mod,
    "Pramukh": run_pramukh_mod,
    "R N Export": run_rn_export_mod,
    "Rapnet":run_rapnet_mod,
    "R Rajesh": run_r_rajesh_mod,
    "Ratnakala Mail": run_ratnakala_mail_mod,
    "Ratnakala Web": run_ratnakala_web_mod,
    "Ratnakala Fancy": run_ratnakala_fancy_mod,
    "Rays Mail": run_rays_mod,
    "Rays Web": run_rays_web_mod,
    "Red Exim": run_red_exim_mod,
    "S Vinod": run_s_vinod_mod,
    "Sagar List": run_sagar_list_mod,
    "Sanghvi & Sons": run_sanghvi_sons_mod,
    "Shairu": run_shairu_mod,
    "Sheetal": run_sheetal_mod,
    "SKE":run_ske_mod,
    "Shivam List": run_shivan_mod,
    "Shyam Copration": run_shyam_mod,
    "SKE": run_ske_mod,
    "SNJ": run_snj_mod,
    "SRD":run_srd_mod,
    "SRK":run_srk_mod,
    "Starlight Diamond":run_starlight_mod,
    "Starrays":run_starrays_mod,
    "Sunrays":run_sunrays_mod,
    "Unique":run_unique_mod,
    "Vaibhav":run_vaibhav_mod,
    "VD Global":run_vd_global_mod,
    "Veer Gems":run_veer_gems_mod,
    "Vinay Diamond":run_vinay_diamonds_mod,
    "Kiran Fancy":run_kiran_fancy_mod,
    "Lakhi Fancy":run_lakhi_fancy_mod
}

REGISTRY = {**DEAL_FUNCTION_MAP, **REPORT_FUNCTION_MAP}

# -------------------------------
# STYLES
# -------------------------------
st.markdown(
    """
    <style>
    /* Hide Streamlit chrome */
    #MainMenu,
    footer,
    header,
    [data-testid="stDecoration"] {
        display: none !important;
        visibility: hidden !important;
    }

    /* Page */
    html,
    body,
    [data-testid="stAppViewContainer"],
    .stApp {
        background-color: #f2f2f2 !important;
        color: #333333 !important;
    }

    .block-container {
        padding-top: 8rem !important;
    }

    /* Logo */
    .center-logo {
        display: flex;
        justify-content: center;
        align-items: center;
        width: 300px;
        max-width: 100%;
        margin: 0 auto 25px auto;
        padding: 12px;
        border-radius: 10px;
        background-color: #333333;
        box-sizing: border-box;
    }

    .center-logo img {
        display: block;
        width: 180px;
        max-width: 100%;
        height: auto;
    }

    .login-title {
        margin-bottom: 20px;
        color: #333333;
        font-size: 42px;
        font-weight: 800;
        text-align: center;
    }

    /* Main login box */
    div[data-testid="stVerticalBlockBorderWrapper"] {
        border: 1px solid #cfcfcf !important;
        border-radius: 8px !important;
    }

    /* Text and password inputs */
    div[data-testid="stTextInput"]
    div[data-testid="stTextInputRootElement"] {
        min-height: 45px !important;
        border: 1px solid #000000 !important;
        border-radius: 7px !important;
        background-color: #ffffff !important;
        box-shadow: none !important;
        box-sizing: border-box !important;
        overflow: hidden !important;
    }

    div[data-testid="stTextInput"]
    div[data-testid="stTextInputRootElement"]:focus-within {
        border: 2px solid #000000 !important;
        box-shadow: none !important;
    }

    div[data-testid="stTextInput"] input {
        padding: 8px 10px !important;
        border: none !important;
        outline: none !important;
        background: transparent !important;
        color: #333333 !important;
        box-shadow: none !important;
    }

    div[data-testid="stTextInput"] input::placeholder {
        color: #6b7280 !important;
        opacity: 1 !important;
    }

    div[data-testid="stTextInput"] button {
        margin-right: 6px !important;
        color: #333333 !important;
    }

    /*
       Selectbox border fix.
       In current Streamlit/BaseWeb markup, the visible field is the direct
       child inside [data-baseweb="select"], not the outer select wrapper.
    */
    div[data-testid="stSelectbox"] div[data-baseweb="select"],
    .st-key-deal_select div[data-baseweb="select"],
    .st-key-report_select div[data-baseweb="select"] {
        min-height: 45px !important;
        border: 0 !important;
        border-radius: 7px !important;
        background-color: transparent !important;
        box-shadow: none !important;
    }

    div[data-testid="stSelectbox"] div[data-baseweb="select"] > div,
    .st-key-deal_select div[data-baseweb="select"] > div,
    .st-key-report_select div[data-baseweb="select"] > div {
        min-height: 45px !important;
        border: 1px solid #000000 !important;
        border-radius: 7px !important;
        outline: none !important;
        background-color: #ffffff !important;
        box-shadow: none !important;
        box-sizing: border-box !important;
    }

    div[data-testid="stSelectbox"] div[data-baseweb="select"]:focus-within > div,
    .st-key-deal_select div[data-baseweb="select"]:focus-within > div,
    .st-key-report_select div[data-baseweb="select"]:focus-within > div {
        border: 2px solid #000000 !important;
        box-shadow: none !important;
    }

    div[data-testid="stSelectbox"] div[data-baseweb="select"] input,
    div[data-testid="stSelectbox"] div[data-baseweb="select"] span,
    div[data-testid="stSelectbox"] div[data-baseweb="select"] svg,
    .st-key-deal_select div[data-baseweb="select"] input,
    .st-key-deal_select div[data-baseweb="select"] span,
    .st-key-deal_select div[data-baseweb="select"] svg,
    .st-key-report_select div[data-baseweb="select"] input,
    .st-key-report_select div[data-baseweb="select"] span,
    .st-key-report_select div[data-baseweb="select"] svg {
        color: #333333 !important;
    }

    div[data-testid="stSelectbox"] div[data-baseweb="select"] svg,
    .st-key-deal_select div[data-baseweb="select"] svg,
    .st-key-report_select div[data-baseweb="select"] svg {
        fill: currentColor !important;
    }

    /* Multiselect border fix */
    div[data-testid="stMultiSelect"] div[data-baseweb="select"] {
        min-height: 45px !important;
        border: 0 !important;
        border-radius: 7px !important;
        background-color: transparent !important;
        box-shadow: none !important;
    }

    div[data-testid="stMultiSelect"] div[data-baseweb="select"] > div {
        min-height: 45px !important;
        border: 1px solid #000000 !important;
        border-radius: 7px !important;
        outline: none !important;
        background-color: #ffffff !important;
        box-shadow: none !important;
        box-sizing: border-box !important;
    }

    div[data-testid="stMultiSelect"] div[data-baseweb="select"]:focus-within > div {
        border: 2px solid #000000 !important;
        box-shadow: none !important;
    }

    div[data-testid="stMultiSelect"] div[data-baseweb="select"] input,
    div[data-testid="stMultiSelect"] div[data-baseweb="select"] span,
    div[data-testid="stMultiSelect"] div[data-baseweb="select"] svg {
        color: #333333 !important;
    }

    div[data-testid="stMultiSelect"] div[data-baseweb="select"] svg {
        fill: currentColor !important;
    }

    /* Dark theme */
    @media (prefers-color-scheme: dark) {
        html,
        body,
        [data-testid="stAppViewContainer"],
        .stApp {
            background-color: #0b0b0b !important;
            color: #f2f2f2 !important;
        }

        .login-title {
            color: #f2f2f2 !important;
        }

        .center-logo {
            background-color: #444444 !important;
        }

        div[data-testid="stVerticalBlockBorderWrapper"] {
            border-color: #555555 !important;
        }

        div[data-testid="stTextInput"]
        div[data-testid="stTextInputRootElement"] {
            border-color: #f2f2f2 !important;
            background-color: #111111 !important;
        }

        div[data-testid="stTextInput"]
        div[data-testid="stTextInputRootElement"]:focus-within {
            border-color: #f2f2f2 !important;
        }

        div[data-testid="stTextInput"] input,
        div[data-testid="stTextInput"] button {
            color: #f2f2f2 !important;
        }

        div[data-testid="stTextInput"] input::placeholder {
            color: #cfcfcf !important;
        }

        div[data-testid="stSelectbox"] div[data-baseweb="select"] > div,
        .st-key-deal_select div[data-baseweb="select"] > div,
        .st-key-report_select div[data-baseweb="select"] > div,
        div[data-testid="stMultiSelect"] div[data-baseweb="select"] > div {
            border-color: #f2f2f2 !important;
            background-color: #111111 !important;
        }

        div[data-testid="stSelectbox"] div[data-baseweb="select"]:focus-within > div,
        .st-key-deal_select div[data-baseweb="select"]:focus-within > div,
        .st-key-report_select div[data-baseweb="select"]:focus-within > div,
        div[data-testid="stMultiSelect"] div[data-baseweb="select"]:focus-within > div {
            border-color: #f2f2f2 !important;
        }

        div[data-testid="stSelectbox"] div[data-baseweb="select"] input,
        div[data-testid="stSelectbox"] div[data-baseweb="select"] span,
        div[data-testid="stSelectbox"] div[data-baseweb="select"] svg,
        .st-key-deal_select div[data-baseweb="select"] input,
        .st-key-deal_select div[data-baseweb="select"] span,
        .st-key-deal_select div[data-baseweb="select"] svg,
        .st-key-report_select div[data-baseweb="select"] input,
        .st-key-report_select div[data-baseweb="select"] span,
        .st-key-report_select div[data-baseweb="select"] svg,
        div[data-testid="stMultiSelect"] div[data-baseweb="select"] input,
        div[data-testid="stMultiSelect"] div[data-baseweb="select"] span,
        div[data-testid="stMultiSelect"] div[data-baseweb="select"] svg {
            color: #f2f2f2 !important;
        }
    }
    </style>
    """,
    unsafe_allow_html=True,
)

# Display the logo in the center
st.markdown(
    f'<div class="center-logo"><img src="{APP_ICON}" width="180" alt="KP Sanghvi Logo"/></div>',
    unsafe_allow_html=True,
)

# -------------------------------
# NAVIGATION HELPERS (session state is the source of truth)
# -------------------------------
MODE_KEY = "mode"

def get_mode(default: str = HOME) -> str:
    # Prefer session state; seed once from URL if missing
    if MODE_KEY not in st.session_state:
        url_mode = st.query_params.get("mode", default)
        st.session_state[MODE_KEY] = url_mode
    return st.session_state[MODE_KEY]

def set_mode(new_mode: str | None):
    # Update session state
    st.session_state[MODE_KEY] = new_mode or HOME
    # Mirror to URL for shareability
    if new_mode is None or new_mode == HOME:
        for k in ("mode", "deal", "report"):
            try:
                st.query_params.pop(k)
            except KeyError:
                pass
    else:
        st.query_params["mode"] = new_mode

def set_qp_bulk(updates: dict):
    """Apply multiple query-param updates; remove keys where value is None."""
    for k, v in updates.items():
        if v is None:
            try:
                st.query_params.pop(k)
            except KeyError:
                pass
        else:
            st.query_params[k] = v

# -------------------------------
# AUTH STATE
# -------------------------------
if "logged_in" not in st.session_state:
    st.session_state.logged_in = False

# -------------------------------
# PAGES
# -------------------------------
def login_page():
    # Initialize session state once
    if "logged_in" not in st.session_state:
        st.session_state.logged_in = False

    if "username" not in st.session_state:
        st.session_state.username = ""

    # Center the login form
    column_login = st.columns([1, 1, 1, 2, 1, 1,1])
    with column_login[3]:

        column_login_tit = st.columns([1, 1, 1])
        with column_login_tit[1]:
            st.title("Login")

        with st.form("login_form"):
            username = st.text_input(
                " ",
                placeholder="Username",
                key="login_user",
            )

            password = st.text_input(
                " ",
                placeholder="Password",
                type="password",
                key="login_pass"
            )
            column_login_btn = st.columns([1, 1, 1, 2, 1, 1, 1])
            with column_login_btn[3]:
                submitted = st.form_submit_button("**Login**", use_container_width=True)

        # JS: Enter on first input -> focus second input
        # Enter on second input -> click login button
        st.iframe(
            """
            <script>
            function setupLoginFlow() {
                const parentDoc = window.parent.document;

                const inputs = parentDoc.querySelectorAll("input");
                const buttons = parentDoc.querySelectorAll("button");

                let usernameInput = null;
                let passwordInput = null;
                let loginButton = null;

                for (const input of inputs) {
                    const placeholder = input.getAttribute("placeholder");

                    if (placeholder === "Username") {
                        usernameInput = input;
                    }

                    if (placeholder === "Password") {
                        passwordInput = input;
                    }
                }

                for (const button of buttons) {
                    if (button.innerText.trim() === "Login") {
                        loginButton = button;
                        break;
                    }
                }

                if (
                    usernameInput &&
                    passwordInput &&
                    usernameInput.dataset.loginFlowBound !== "true"
                ) {
                    usernameInput.dataset.loginFlowBound = "true";

                    usernameInput.addEventListener("keydown", function (event) {
                        if (event.key === "Enter") {
                            event.preventDefault();
                            passwordInput.focus();
                        }
                    });
                }

                if (
                    passwordInput &&
                    loginButton &&
                    passwordInput.dataset.loginFlowBound !== "true"
                ) {
                    passwordInput.dataset.loginFlowBound = "true";

                    passwordInput.addEventListener("keydown", function (event) {
                        if (event.key === "Enter") {
                            event.preventDefault();
                            loginButton.click();
                        }
                    });
                }
            }

            setTimeout(setupLoginFlow, 500);
            </script>
            """,
            height=1,
            tab_index=-1,
        )
        
        if submitted:
            if username == USERNAME and password == PASSWORD:
                st.session_state.logged_in = True
                st.session_state.username = username
                st.success("Logged in successfully!")
                st.rerun()
            else:
                st.error("Invalid username or password")

def home_page():
    mode = get_mode(HOME)
    deal_selected = st.query_params.get("deal", SENTINEL)
    report_selected = st.query_params.get("report", SENTINEL)

    deal_options = [SENTINEL] + list(DEAL_FUNCTION_MAP.keys())
    report_options = [SENTINEL] + list(REPORT_FUNCTION_MAP.keys())

    # Top navigation / Home
    if mode != HOME:
        if st.button("🏠 Home", key="btn_home_top"):
            set_qp_bulk({"mode": HOME, "deal": None, "report": None})
            set_mode(HOME)
            st.rerun()
    else:
        # Logout button (Top Right)
        cols = st.columns([1,1,1,1,1,1,1,1,1,1,1])
        with cols[-1]:
            if st.button("Logout", key="btn_logout"):
                st.session_state.logged_in = False
                st.session_state.username = None
                set_qp_bulk({"mode": HOME, "deal": None, "report": None})
                set_mode(HOME)
                st.rerun()

        col1, col2, col3, col4, col5, col6, col7, col8 = st.columns(8)

        # Home button (redundant on Home but kept for consistency)
        with col1:
            if st.button("🏠 Home", key="btn_home_main"):
                set_qp_bulk({"mode": HOME, "deal": None, "report": None})
                set_mode(HOME)
                st.rerun()
                
        with col2:
            st.markdown("### Files")
            if st.button("🧾 AHS", key="btn_ahs"):
                set_qp_bulk({"mode": "🧾 AHS", "deal": None, "report": None})
                set_mode("🧾 AHS")
                st.rerun()

            if st.button("🧾 Canadian", key="btn_canadian"):
                set_qp_bulk({"mode": "🧾 Canadian", "deal": None, "report": None})
                set_mode("🧾 Canadian")
                st.rerun()

            if st.button("🧾 LEDS", key="btn_leds"):
                set_qp_bulk({"mode": "🧾 LEDS", "deal": None, "report": None})
                set_mode("🧾 LEDS")
                st.rerun()

            if st.button("🧾 MPB Bill", key="btn_mpb"):
                set_qp_bulk({"mode": "🧾 MPB Bill", "deal": None, "report": None})
                set_mode("🧾 MPB Bill")
                st.rerun()

            if st.button(NAV_FEEDING, key="btn_feeding"):
                set_qp_bulk({"mode": NAV_FEEDING, "deal": None, "report": None})
                set_mode(NAV_FEEDING)
                st.rerun()

            if st.button(NAV_MERGE, key="btn_merge"):
                set_qp_bulk({"mode": NAV_MERGE, "deal": None, "report": None})
                set_mode(NAV_MERGE)
                st.rerun()

            if st.button(NAV_VLOOKUP, key="btn_vlookup"):
                set_qp_bulk({"mode": NAV_VLOOKUP, "deal": None, "report": None})
                set_mode(NAV_VLOOKUP)
                st.rerun()

        with col3:
            st.markdown("### Stock File")
            if st.button("🧾 Convert Stock List", key="btn_convert"):
                set_qp_bulk({"mode": "🧾 Convert Stock List", "deal": None, "report": None})
                set_mode("🧾 Convert Stock List")
                st.rerun()
           
            if st.button("🧾 Browns File", key="btn_brown"):
                set_qp_bulk({"mode": "🧾 Browns File", "deal": None, "report": None})
                set_mode("🧾 Browns File")
                st.rerun()

            if st.button("🧾 Pending Stock", key="btn_pending"):
                set_qp_bulk({"mode": "🧾 Pending Stock", "deal": None, "report": None})
                set_mode("🧾 Pending Stock")
                st.rerun()

            if st.button("🧾 GIA Price", key="btn_gia_price"):
                set_qp_bulk({"mode": "🧾 GIA Price", "deal": None, "report": None})
                set_mode("🧾 GIA Price")
                st.rerun()
                
            if st.button("🧾 Market Disc", key="btn_market_disc"):
                set_qp_bulk({"mode": "🧾 Market Disc", "deal": None, "report": None})
                set_mode("🧾 Market Disc")
                st.rerun()
            
            if st.button("🧾 MPB Nattes ", key="btn_mpb_ntts_disc"):
                set_qp_bulk({"mode": "🧾 MPB Nattes ", "deal": None, "report": None})
                set_mode("🧾 MPB Nattes ")
                st.rerun()

        with col4:
            st.markdown("### Premal Sir")
            if st.button("🧾 New Stone", key="btn_new_stone"):
                set_qp_bulk({"mode": "🧾 New Stone", "deal": None, "report": None})
                set_mode("🧾 New Stone")
                st.rerun()
            if st.button("🧾 Hemil Check", key="btn_hemil"):
                set_qp_bulk({"mode": "🧾 Hemil Check", "deal": None, "report": None})
                set_mode("🧾 Hemil Check")
                st.rerun()
            
            if st.button("🧾 Revision File", key="btn_revision"):
                set_qp_bulk({"mode": "🧾 Revision File", "deal": None, "report": None})
                set_mode("🧾 Revision File")
                st.rerun()
            
            if st.button("🧾 PZ Revision Stock File", key="btn_revision_stock"):
                set_qp_bulk({"mode": "🧾 PZ Revision Stock File", "deal": None, "report": None})
                set_mode("🧾 PZ Revision Stock File")
                st.rerun()

        with col5:
            st.markdown("### Rapnet Files")
            if st.button("🧾 Rapnet Format", key="btn_rapnet"):
                set_qp_bulk({"mode": "🧾 Rapnet Format", "deal": None, "report": None})
                set_mode("🧾 Rapnet Format")
                st.rerun()

            if st.button("🧾 Highlight Stone", key="btn_highlight"):
                set_qp_bulk({"mode": "🧾 Highlight Stone", "deal": None, "report": None})
                set_mode("🧾 Highlight Stone")
                st.rerun()
            
            if st.button("🧾 Rap report", key="btn_rapreport"):
                set_qp_bulk({"mode": "🧾 Rap report", "deal": None, "report": None})
                set_mode("🧾 Rap report")
                st.rerun()
            
            if st.button("🧾 Rap report PDF To Excel", key="btn_rapreport_pdf"):
                set_qp_bulk({"mode": "🧾 Rap report PDF To Excel", "deal": None, "report": None})
                set_mode("🧾 Rap report PDF To Excel")
                st.rerun()
            
            if st.button("🧾 Dossier Report", key="btn_dossier"):
                set_qp_bulk({"mode": "🧾 Dossier Report", "deal": None, "report": None})
                set_mode("🧾 Dossier Report")
                st.rerun()
            if st.button("🧾 Dossier report PDF To Excel", key="btn_dossier_pdf"):
                set_qp_bulk({"mode": "🧾 Dossier report PDF To Excel", "deal": None, "report": None})
                set_mode("🧾 Dossier report PDF To Excel")
                st.rerun()
            
            if st.button("🧾 Sale File", key="btn_sale_file"):
                set_qp_bulk({"mode": "🧾 Sale File", "deal": None, "report": None})
                set_mode("🧾 Sale File")
                st.rerun()

                # PDF To EXCEL Files
        with col6:
            st.markdown("### PDF To Excel")

            if st.button("🧾 Modify HRD PDF"):
                set_qp_bulk({"mode": "🧾 Modify HRD PDF", "deal": None, "report": None})
                set_mode("🧾 Modify HRD PDF")
                st.rerun()
            
            if st.button("🧾 Print Sticker", key="btn_print_sticker"):
                set_qp_bulk({"mode": "🧾 Print Sticker", "deal": None, "report": None})
                set_mode("🧾 Print Sticker")
                st.rerun()

        with col7:
            st.markdown("### Export Files ")

            if st.button("🧾 Export File", key="btn_export"):
                set_qp_bulk({"mode": "🧾 Export File", "deal": None, "report": None})
                set_mode("🧾 Export File")
                st.rerun()

            if st.button("🧾 Black Lookup Sheet", key="btn_all_black"):
                set_qp_bulk({"mode": "🧾 Black Lookup Sheet", "deal": None, "report": None})
                set_mode("🧾 Black Lookup Sheet")
                st.rerun()

            if st.button("🧾 Websit File Format", key="btn_webformat"):
                set_qp_bulk({"mode": "🧾 Websit File Format", "deal": None, "report": None})
                set_mode("🧾 Websit File Format")
                st.rerun()

        with col8:
            st.markdown("### DEAL Files")
            # Deal select
            deal_index = deal_options.index(deal_selected) if deal_selected in deal_options else 0
            deal_value = st.selectbox(
                "Deal Files:", options=deal_options, index=deal_index, key="deal_select"
            )
            if deal_value != deal_selected:
                if deal_value == SENTINEL:
                    set_qp_bulk({"deal": None})
                else:
                    set_qp_bulk({"deal": deal_value})
                    set_mode(deal_value)
                st.rerun()

            # Report select
            report_index = report_options.index(report_selected) if report_selected in report_options else 0
            report_value = st.selectbox(
                "Black Vlookup Files:", options=report_options, index=report_index, key="report_select"
            )
            if report_value != report_selected:
                if report_value == SENTINEL:
                    set_qp_bulk({"report": None})
                else:
                    set_qp_bulk({"report": report_value})
                    set_mode(report_value)
                st.rerun()

        with col7:
            st.empty()

    # -------------------------------
    # DISPATCHER
    # -------------------------------
    if mode == HOME:
        st.markdown("")
    elif mode == NAV_VLOOKUP:
        run_vlookup_mod()
    elif mode == NAV_MERGE:
        run_merge_mod()
    elif mode == NAV_FEEDING:
        run_feeding_mod()
    elif mode in REGISTRY:
        REGISTRY[mode]()
    elif mode == "🧾 AHS":
        run_ahs_mod()
    elif mode == "🧾 Canadian":
        run_canadian_mod()
    elif mode == "🧾 LEDS":
        run_led_mod()
    elif mode == "🧾 MPB Bill":
        run_mpb_bill()
    elif mode == "🧾 Market Disc":
        run_market_disc()
    elif mode == "🧾 Revision File":
        run_revsion()
    elif mode == "🧾 New Stone":
        run_stock_list()
    elif mode == "🧾 Hemil Check":
        run_hemil_check()
    elif mode == "🧾 Convert Stock List":
        run_convert_stock_mod()
    elif mode == "🧾 Rapnet Format":
        run_rapnet_format_mod()
    elif mode == "🧾 Pending Stock":
        run_pending_stock()
    elif mode == "🧾 GIA Price":
        run_gia_price()
    elif mode == "🧾 Highlight Stone":
        run_highlight()
    elif mode == "🧾 Rap report PDF To Excel":
        rap_report_pdf()
    elif mode == "🧾 Rap report":
        rap_report()
    elif mode == "🧾 Dossier report PDF To Excel":
        dosser_run()
    elif mode == "🧾 Dossier Report":
        dossier_report()
    elif mode == "🧾 Sale File":
        sale_file()
    elif mode == "🧾 Modify HRD PDF":
        pdf_modify_run()
    elif mode == "🧾 Export File":
        run_export()
    elif mode == "🧾 Black Lookup Sheet":
        run_black_mod()
    elif mode == "🧾 Websit File Format":
        run_web_fromat()
    elif mode == "🧾 Browns File":
        run_browns_mod()
    elif mode == "🧾 MPB Nattes ":
        run_mpb_natts_mod()
    elif mode == "🧾 Print Sticker":
        run_print_sticker()
    elif mode == "🧾 PZ Revision Stock File":
        revision_stock_list()
    
    else:
        st.warning("Unknown mode. Returning Home…")
        set_qp_bulk({"mode": HOME, "deal": None, "report": None})
        set_mode(HOME)
        st.rerun()

if st.session_state.logged_in:
    home_page()
else:
    login_page()
