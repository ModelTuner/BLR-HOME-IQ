import streamlit as st
import pandas as pd
import joblib
import base64
import html as html_lib
import json
import time
import math
import importlib.util
from io import BytesIO
from datetime import datetime
from pathlib import Path

try:
    from supabase import create_client
    SUPABASE_AVAILABLE = True
except ImportError:
    SUPABASE_AVAILABLE = False

EXCEL_AVAILABLE = importlib.util.find_spec("openpyxl") is not None
if EXCEL_AVAILABLE:
    try:
        # Resolve these dynamically so environments with incomplete openpyxl
        # type/source metadata do not flag direct imports as unresolved.
        _openpyxl_styles = importlib.import_module("openpyxl.styles")
        _openpyxl_utils = importlib.import_module("openpyxl.utils")
        Alignment = _openpyxl_styles.Alignment
        Border = _openpyxl_styles.Border
        Font = _openpyxl_styles.Font
        PatternFill = _openpyxl_styles.PatternFill
        Side = _openpyxl_styles.Side
        get_column_letter = _openpyxl_utils.get_column_letter
    except ImportError:
        EXCEL_AVAILABLE = False

st.set_page_config(
    page_title="BLR HOME IQ | Bengaluru Property Intelligence",
    page_icon="🏡",
    layout="wide",
    initial_sidebar_state="expanded",
)

# -------------------- MODEL PATHS --------------------
BASE_DIR = Path(__file__).resolve().parent
INDIA_MODEL_PATH = str(BASE_DIR / "india_house_price_model.pkl")
# Prefer the finalized dataset; retain compatibility with the previous filename.
DATA_CANDIDATES = [
    BASE_DIR / "DATA" / "Bengaluru_House_Data_Final.csv",
    BASE_DIR / "DATA" / "Bengaluru_House_Data_Cleaned.csv",
]
INDIA_DATA_PATH = str(next((p for p in DATA_CANDIDATES if p.is_file()), DATA_CANDIDATES[0]))
HISTORY_COLUMNS = [
    "Date and time", "Country", "Estimated price", "Currency", "Property details"
]

# Soft input check only (a warning, never blocks the estimate).
MIN_SQFT_PER_BHK = 250

# Dashboard pages (sidebar navigation)
PAGES = ("home", "predict", "history", "profile")
PAGE_LABEL_KEYS = {
    "home": "nav_home",
    "predict": "nav_predict",
    "history": "nav_history",
    "profile": "nav_profile",
}
PAGE_ICONS = {
    "home": ":material/home:",
    "predict": ":material/calculate:",
    "history": ":material/history:",
    "profile": ":material/person:",
}


# ==================== LANGUAGE (Indian languages) ====================
# Saved CSV files always stay in English; selected interface labels change language. Missing keys safely fall back to English.
TEXT = {
    "English": {
        # general
        "select_language": "🌐 Language / भाषा",
        "country_India": "Bengaluru",
        "brand_sub": "Bengaluru Property Intelligence",
        "hero_eyebrow": "AI-POWERED VALUATION",
        "hero_welcome": "Welcome to",
        "hero_title": "Find the value<br>behind every home.",
        "hero_text": "Explore an intelligent estimate built from property characteristics. Add the details below to generate your valuation.",
        "hero_tag1": "✦ Machine learning",
        "hero_tag2": "⌂ Property insights",
        "hero_tag3": "↗ Instant estimate",
        "configure_property": "Configure your property",
        "your_valuation": "Your valuation",
        "submit_hint": "Submit your property details to see the estimate.",
        "estimate_btn": "✦ Estimate property value",
        "unit_lakh": "lakh",
        "unit_crore": "crore",
        # navigation
        "nav_home": "Home",
        "nav_predict": "Predict Property Value",
        "nav_history": "My History",
        "nav_profile": "Profile",
        "nav_tagline": "Smart property. Better decisions.",
        # India form
        "india_sub": "Bengaluru property estimate · Prices in ₹ lakh / crore",
        "in_location": "Location (Bengaluru)",
        "in_area_type": "Property type",
        "in_bhk": "Bedrooms / BHK",
        "in_area": "Built-up area (sq ft)",
        "in_bath": "Bathrooms",
        "in_balcony": "Balconies",
        "sec_location": "Location and property type",
        "sec_location_sub": "Where is the property, and which area type does the listing show?",
        "sec_size": "Size and rooms",
        "sec_size_sub": "Built-up area, bedrooms, bathrooms and balconies.",
        "help_location": "Choose the Bengaluru locality of the property.",
        "help_area_type": "Pick the area type exactly as shown in the property listing.",
        "help_bhk": "BHK = number of bedrooms (with a hall and kitchen).",
        "help_area": "Total built-up area in square feet.",
        "help_bath": "Number of bathrooms.",
        "help_balcony": "Number of balconies (0 if there is none).",
        "spinner_predict": "Calculating your estimate...",
        "warn_small_area": "{area} sq ft looks small for {bhk} BHK. The estimate may be less reliable, so please re-check the area.",
        # messages
        "err_prediction": "Prediction could not be generated. Check the model files and inputs.",
        "err_model_load": "The model or dataset file could not be loaded ({name}). Keep it in the same folder as app.py.",
        "tech_details": "Technical details",
        # result panel
        "result_label": "Estimated sale price",
        "result_ready": "Ready when you are",
        "result_ready_note": "Complete the form and estimate the value.",
        "step1": "Choose the location and property type",
        "step2": "Enter area, BHK, bathrooms and balconies",
        "step3": "Press the estimate button to see the price",
        "note_india": "Model estimate · INR · not guaranteed",
        "model_extra_trees": "Extra Trees",
        "dataset_india": "Bengaluru · India",
        "inputs_india": "6 property inputs",
        "stat_model": "Model",
        "stat_inputs": "Inputs",
        "stat_dataset": "Dataset",
        "stat_location": "Location",
        "stat_rate": "Approx. rate per sq ft",
        "summary_title": "Property summary",
        "last_details": "View details used for your last prediction",
        # home cards
        "card_details_title": "View details",
        "card_details_text": "See the exact inputs used for your last prediction.",
        "card_history_title": "Prediction history",
        "card_history_text": "Open your saved predictions and manage your data.",
        "card_download_title": "Download history",
        "card_download_text": "Get your saved history in Excel or CSV format.",
        "open_btn": "Open",
        "languages_title": "Indian languages",
        "languages_sub": "Change the language from the menu. Your typed inputs are kept.",
        # profile
        "profile_title": "Your profile",
        "profile_sub": "Your account and information about this app.",
        "profile_email": "Email",
        "profile_language": "Language",
        "profile_about_title": "About this estimate",
        "profile_about_text": "BLR HOME IQ uses a machine-learning model trained on a Bengaluru property dataset. It works only for Bengaluru localities and gives an educational estimate. It is not a valuation certificate or financial advice.",
        # history
        "history_title": "{country} prediction history",
        "history_sub": "Only your {country} predictions are shown here. They are saved in your account and only you can see them.",
        "history_empty_title": "No saved predictions yet",
        "history_empty": "No saved {country} predictions yet. Submit a property estimate to start your history.",
        "go_predict": "Estimate a property",
        "history_count": "Saved {country} predictions: {count}",
        "download_excel": "⬇️ Download {country} history (Excel)",
        "download_csv": "⬇️ Download {country} history (CSV)",
        "excel_hint": "For a formatted Excel download, run: pip install openpyxl",
        "clear_btn": "🗑️ Clear {country} history",
        "history_cleared": "{country} prediction history cleared.",
        "sheet_name": "{country} history",
        # history table columns
        "col_datetime": "Date and time",
        "col_price": "Estimated price",
        "col_location": "Location",
        "col_area_type": "Property type",
        "col_bhk": "BHK",
        "col_area": "Area (sq ft)",
        "col_bath": "Bathrooms",
        "col_balcony": "Balconies",
        # footer
        "footer_left": "BLR HOME IQ · Bengaluru Property Intelligence",
        "footer_right": "Educational estimate only · Bengaluru dataset · INR",
        # login / sign up
        "login_title": "Welcome to BLR HOME IQ",
        "login_sub": "Log in or create a free account to continue.",
        "tab_login": "Log in",
        "tab_signup": "Sign up",
        "email": "Email address",
        "password": "Password",
        "confirm_password": "Confirm password",
        "login_btn": "Log in",
        "signup_btn": "Create account",
        "logout_btn": "Log out",
        "logged_in_as": "Logged in as {email}",
        "welcome_user": "Welcome",
        "err_secrets": "Login service is not set up. Check .streamlit/secrets.toml (SUPABASE_URL and SUPABASE_KEY).",
        "err_no_supabase": "The 'supabase' package is not installed. Run: pip install supabase",
        "err_email_invalid": "Please enter a valid email address.",
        "err_pw_short": "Password must be at least 8 characters.",
        "err_pw_mismatch": "Passwords do not match.",
        "err_invalid_login": "Wrong email or password.",
        "err_email_not_confirmed": "Please confirm your email first. Check your inbox for the confirmation link.",
        "err_email_exists": "This email is already registered. Please log in.",
        "err_rate_limit": "Too many attempts or emails sent. Please wait a few minutes and try again.",
        "err_network": "Could not reach the login service. Check your internet connection and try again.",
        "err_generic": "Something went wrong: {error}",
        "signup_check_email": "Account created. Please check your email and click the confirmation link, then log in.",
        "err_session_expired": "Your session has expired. Please log in again.",
        "err_history_load": "Could not load your history: {error}",
        "err_history_clear": "Could not clear your history: {error}",
        "warn_history_save": "Your estimate is shown below, but it could not be saved to your history: {error}",
    },
    "हिन्दी": {
        # general
        "select_language": "🌐 Language / भाषा",
        "country_India": "बेंगलुरु",
        "brand_sub": "प्रॉपर्टी इंटेलिजेंस",
        "hero_eyebrow": "AI-संचालित मूल्यांकन",
        "hero_welcome": "स्वागत है",
        "hero_title": "हर घर की कीमत<br>जानिए, आसानी से।",
        "hero_text": "प्रॉपर्टी की विशेषताओं के आधार पर बना स्मार्ट अनुमान देखें। अपना मूल्यांकन पाने के लिए नीचे विवरण भरें।",
        "hero_tag1": "✦ मशीन लर्निंग",
        "hero_tag2": "⌂ प्रॉपर्टी इनसाइट्स",
        "hero_tag3": "↗ तुरंत अनुमान",
        "configure_property": "अपनी प्रॉपर्टी की जानकारी भरें",
        "your_valuation": "आपका मूल्यांकन",
        "submit_hint": "अनुमान देखने के लिए अपनी प्रॉपर्टी का विवरण भरें।",
        "estimate_btn": "✦ प्रॉपर्टी की कीमत का अनुमान लगाएँ",
        "unit_lakh": "लाख",
        "unit_crore": "करोड़",
        # navigation
        "nav_home": "होम",
        "nav_predict": "प्रॉपर्टी की कीमत जानें",
        "nav_history": "मेरा इतिहास",
        "nav_profile": "प्रोफ़ाइल",
        "nav_tagline": "स्मार्ट प्रॉपर्टी। बेहतर फ़ैसले।",
        # India form
        "india_sub": "बेंगलुरु प्रॉपर्टी का अनुमान · कीमतें ₹ लाख / करोड़ में",
        "in_location": "स्थान (बेंगलुरु)",
        "in_area_type": "प्रॉपर्टी का प्रकार",
        "in_bhk": "बेडरूम / BHK",
        "in_area": "बिल्ट-अप क्षेत्रफल (वर्ग फुट)",
        "in_bath": "बाथरूम",
        "in_balcony": "बालकनी",
        "sec_location": "स्थान और प्रॉपर्टी का प्रकार",
        "sec_location_sub": "प्रॉपर्टी कहाँ है और लिस्टिंग में कौन-सा क्षेत्रफल प्रकार दिया है?",
        "sec_size": "आकार और कमरे",
        "sec_size_sub": "बिल्ट-अप क्षेत्रफल, बेडरूम, बाथरूम और बालकनी।",
        "help_location": "प्रॉपर्टी का बेंगलुरु का इलाका चुनें।",
        "help_area_type": "प्रॉपर्टी लिस्टिंग में जैसा लिखा है, वैसा क्षेत्रफल प्रकार चुनें।",
        "help_bhk": "BHK यानी बेडरूम की संख्या (हॉल और किचन के साथ)।",
        "help_area": "कुल बिल्ट-अप क्षेत्रफल, वर्ग फुट में।",
        "help_bath": "बाथरूम की संख्या।",
        "help_balcony": "बालकनी की संख्या (नहीं है तो 0)।",
        "spinner_predict": "आपका अनुमान तैयार हो रहा है...",
        "warn_small_area": "{bhk} BHK के लिए {area} वर्ग फुट कम लगता है। अनुमान कम भरोसेमंद हो सकता है, कृपया क्षेत्रफल दोबारा जाँचें।",
        # messages
        "err_prediction": "अनुमान तैयार नहीं हो सका। मॉडल फ़ाइलें और इनपुट जाँचें।",
        "err_model_load": "मॉडल या डेटासेट फ़ाइल लोड नहीं हो सकी ({name})। इसे app.py वाले फ़ोल्डर में रखें।",
        "tech_details": "तकनीकी जानकारी",
        # result panel
        "result_label": "अनुमानित बिक्री मूल्य",
        "result_ready": "आप तैयार हों तो शुरू करें",
        "result_ready_note": "फ़ॉर्म भरें और कीमत का अनुमान लगाएँ।",
        "step1": "स्थान और प्रॉपर्टी का प्रकार चुनें",
        "step2": "क्षेत्रफल, BHK, बाथरूम और बालकनी भरें",
        "step3": "कीमत देखने के लिए अनुमान बटन दबाएँ",
        "note_india": "मॉडल का अनुमान · INR (₹) · गारंटीशुदा नहीं",
        "model_extra_trees": "एक्स्ट्रा ट्रीज़",
        "dataset_india": "बेंगलुरु · भारत",
        "inputs_india": "6 प्रॉपर्टी इनपुट",
        "stat_model": "मॉडल",
        "stat_inputs": "इनपुट",
        "stat_dataset": "डेटासेट",
        "stat_location": "स्थान",
        "stat_rate": "लगभग दर प्रति वर्ग फुट",
        "summary_title": "प्रॉपर्टी का सार",
        "last_details": "पिछले अनुमान में इस्तेमाल हुई जानकारी देखें",
        # home cards
        "card_details_title": "विवरण देखें",
        "card_details_text": "पिछले अनुमान में इस्तेमाल हुई जानकारी देखें।",
        "card_history_title": "अनुमान इतिहास",
        "card_history_text": "सेव किए गए अनुमान खोलें और अपना डेटा सँभालें।",
        "card_download_title": "इतिहास डाउनलोड",
        "card_download_text": "सेव इतिहास Excel या CSV में पाएँ।",
        "open_btn": "खोलें",
        "languages_title": "भारतीय भाषाएँ",
        "languages_sub": "मेन्यू से भाषा बदलें। आपके भरे हुए इनपुट बने रहते हैं।",
        # profile
        "profile_title": "आपकी प्रोफ़ाइल",
        "profile_sub": "आपका अकाउंट और इस ऐप की जानकारी।",
        "profile_email": "ईमेल",
        "profile_language": "भाषा",
        "profile_about_title": "इस अनुमान के बारे में",
        "profile_about_text": "BLR HOME IQ बेंगलुरु की प्रॉपर्टी के डेटा पर प्रशिक्षित मशीन-लर्निंग मॉडल इस्तेमाल करता है। यह सिर्फ़ बेंगलुरु के इलाकों के लिए है और शैक्षिक अनुमान देता है। यह वैल्यूएशन सर्टिफ़िकेट या वित्तीय सलाह नहीं है।",
        # history
        "history_title": "{country} अनुमान इतिहास",
        "history_sub": "यहाँ सिर्फ़ आपके {country} के अनुमान दिखते हैं। ये आपके अकाउंट में सेव रहते हैं और सिर्फ़ आप इन्हें देख सकते हैं।",
        "history_empty_title": "अभी कोई अनुमान सेव नहीं है",
        "history_empty": "अभी तक {country} का कोई अनुमान सेव नहीं है। इतिहास शुरू करने के लिए प्रॉपर्टी का अनुमान लगाएँ।",
        "go_predict": "प्रॉपर्टी का अनुमान लगाएँ",
        "history_count": "सेव किए गए {country} अनुमान: {count}",
        "download_excel": "⬇️ {country} इतिहास डाउनलोड करें (Excel)",
        "download_csv": "⬇️ {country} इतिहास डाउनलोड करें (CSV)",
        "excel_hint": "सुंदर Excel डाउनलोड के लिए चलाएँ: pip install openpyxl",
        "clear_btn": "🗑️ {country} इतिहास मिटाएँ",
        "history_cleared": "{country} का अनुमान इतिहास मिटा दिया गया।",
        "sheet_name": "{country} इतिहास",
        # history table columns
        "col_datetime": "तारीख और समय",
        "col_price": "अनुमानित कीमत",
        "col_location": "स्थान",
        "col_area_type": "प्रॉपर्टी का प्रकार",
        "col_bhk": "BHK",
        "col_area": "क्षेत्रफल (वर्ग फुट)",
        "col_bath": "बाथरूम",
        "col_balcony": "बालकनी",
        # footer
        "footer_left": "BLR HOME IQ · बेंगलुरु प्रॉपर्टी इंटेलिजेंस",
        "footer_right": "सिर्फ़ शैक्षिक अनुमान · बेंगलुरु डेटासेट · INR",
        # login / sign up
        "login_title": "BLR HOME IQ में आपका स्वागत है",
        "login_sub": "आगे बढ़ने के लिए लॉग इन करें या मुफ़्त अकाउंट बनाएँ।",
        "tab_login": "लॉग इन",
        "tab_signup": "साइन अप",
        "email": "ईमेल पता",
        "password": "पासवर्ड",
        "confirm_password": "पासवर्ड दोबारा लिखें",
        "login_btn": "लॉग इन करें",
        "signup_btn": "अकाउंट बनाएँ",
        "logout_btn": "लॉग आउट",
        "logged_in_as": "लॉग इन: {email}",
        "welcome_user": "स्वागत है",
        "err_secrets": "लॉगिन सेवा सेट नहीं है। .streamlit/secrets.toml (SUPABASE_URL और SUPABASE_KEY) जाँचें।",
        "err_no_supabase": "'supabase' पैकेज इंस्टॉल नहीं है। चलाएँ: pip install supabase",
        "err_email_invalid": "कृपया सही ईमेल पता लिखें।",
        "err_pw_short": "पासवर्ड कम से कम 8 अक्षरों का होना चाहिए।",
        "err_pw_mismatch": "दोनों पासवर्ड एक जैसे नहीं हैं।",
        "err_invalid_login": "ईमेल या पासवर्ड गलत है।",
        "err_email_not_confirmed": "पहले अपना ईमेल कन्फ़र्म करें। इनबॉक्स में भेजा गया लिंक देखें।",
        "err_email_exists": "यह ईमेल पहले से रजिस्टर है। कृपया लॉग इन करें।",
        "err_rate_limit": "बहुत ज़्यादा कोशिशें या ईमेल हो गए। कुछ मिनट रुककर दोबारा कोशिश करें।",
        "err_network": "लॉगिन सेवा से संपर्क नहीं हो सका। इंटरनेट जाँचें और दोबारा कोशिश करें।",
        "err_generic": "कुछ गड़बड़ हो गई: {error}",
        "signup_check_email": "अकाउंट बन गया। कृपया अपना ईमेल देखें, कन्फ़र्मेशन लिंक पर क्लिक करें, फिर लॉग इन करें।",
        "err_session_expired": "आपका सेशन खत्म हो गया है। कृपया दोबारा लॉग इन करें।",
        "err_history_load": "आपका इतिहास लोड नहीं हो सका: {error}",
        "err_history_clear": "आपका इतिहास मिटाया नहीं जा सका: {error}",
        "warn_history_save": "आपका अनुमान नीचे दिख रहा है, लेकिन वह आपके इतिहास में सेव नहीं हो सका: {error}",
    },
    "ਪੰਜਾਬੀ": {
        "select_language": "🌐 ਭਾਸ਼ਾ ਚੁਣੋ", "country_India": "ਬੈਂਗਲੁਰੂ",
        "brand_sub": "ਬੈਂਗਲੁਰੂ ਪ੍ਰਾਪਰਟੀ ਇੰਟੈਲੀਜੈਂਸ",
        "hero_eyebrow": "AI ਨਾਲ ਕੀਮਤ ਦਾ ਅੰਦਾਜ਼ਾ",
        "hero_title": "ਹਰ ਘਰ ਦੀ ਕੀਮਤ<br>ਜਾਣੋ ਆਸਾਨੀ ਨਾਲ।",
        "hero_text": "ਘਰ ਦੀਆਂ ਵਿਸ਼ੇਸ਼ਤਾਵਾਂ ਦੇ ਆਧਾਰ 'ਤੇ ਕੀਮਤ ਦਾ ਅੰਦਾਜ਼ਾ ਲਵੋ।",
        "nav_home": "ਹੋਮ", "nav_predict": "ਪ੍ਰਾਪਰਟੀ ਦੀ ਕੀਮਤ",
        "nav_history": "ਮੇਰਾ ਇਤਿਹਾਸ", "nav_profile": "ਪ੍ਰੋਫਾਈਲ",
        "configure_property": "ਆਪਣੀ ਪ੍ਰਾਪਰਟੀ ਦੀ ਜਾਣਕਾਰੀ ਭਰੋ", "your_valuation": "ਤੁਹਾਡਾ ਅੰਦਾਜ਼ਾ",
        "submit_hint": "ਕੀਮਤ ਦੇਖਣ ਲਈ ਆਪਣੀ ਪ੍ਰਾਪਰਟੀ ਦੀ ਜਾਣਕਾਰੀ ਭਰੋ।",
        "estimate_btn": "✦ ਪ੍ਰਾਪਰਟੀ ਦੀ ਕੀਮਤ ਦਾ ਅੰਦਾਜ਼ਾ ਲਵੋ",
        "in_location": "ਥਾਂ (ਬੈਂਗਲੁਰੂ)", "in_area_type": "ਪ੍ਰਾਪਰਟੀ ਦੀ ਕਿਸਮ",
        "in_bhk": "ਬੈੱਡਰੂਮ / BHK", "in_area": "ਬਿਲਟ-ਅੱਪ ਖੇਤਰ (ਵਰਗ ਫੁੱਟ)",
        "in_bath": "ਬਾਥਰੂਮ", "in_balcony": "ਬਾਲਕਨੀ",
        "result_label": "ਅੰਦਾਜ਼ਿਤ ਵਿਕਰੀ ਕੀਮਤ", "result_ready": "ਤੁਹਾਡੇ ਲਈ ਤਿਆਰ",
        "result_ready_note": "ਫਾਰਮ ਭਰੋ ਅਤੇ ਕੀਮਤ ਦਾ ਅੰਦਾਜ਼ਾ ਲਵੋ।",
        "login_title": "BLR HOME IQ ਵਿੱਚ ਤੁਹਾਡਾ ਸਵਾਗਤ ਹੈ", "tab_login": "ਲਾਗ ਇਨ",
        "tab_signup": "ਸਾਈਨ ਅੱਪ", "email": "ਈਮੇਲ ਪਤਾ", "password": "ਪਾਸਵਰਡ",
        "confirm_password": "ਪਾਸਵਰਡ ਦੀ ਪੁਸ਼ਟੀ ਕਰੋ", "login_btn": "ਲਾਗ ਇਨ",
        "signup_btn": "ਖਾਤਾ ਬਣਾਓ", "logout_btn": "ਲਾਗ ਆਉਟ",
        "footer_left": "BLR HOME IQ · ਬੈਂਗਲੁਰੂ ਪ੍ਰਾਪਰਟੀ ਇੰਟੈਲੀਜੈਂਸ",
    },
    "தமிழ்": {
        "select_language": "🌐 மொழியைத் தேர்ந்தெடுக்கவும்", "country_India": "பெங்களூரு",
        "brand_sub": "பெங்களூரு சொத்து நுண்ணறிவு", "hero_eyebrow": "AI மதிப்பீடு",
        "hero_title": "ஒவ்வொரு வீட்டின்<br>மதிப்பையும் அறியுங்கள்.",
        "hero_text": "சொத்து விவரங்களின் அடிப்படையில் விலை மதிப்பீட்டைப் பெறுங்கள்.",
        "nav_home": "முகப்பு", "nav_predict": "சொத்து விலை மதிப்பீடு",
        "nav_history": "என் வரலாறு", "nav_profile": "சுயவிவரம்",
        "configure_property": "சொத்து விவரங்களை உள்ளிடவும்", "your_valuation": "உங்கள் மதிப்பீடு",
        "submit_hint": "மதிப்பீட்டைப் பார்க்க சொத்து விவரங்களை உள்ளிடவும்.",
        "estimate_btn": "✦ சொத்து விலையை மதிப்பிடு", "in_location": "இடம் (பெங்களூரு)",
        "in_area_type": "சொத்து வகை", "in_bhk": "படுக்கையறைகள் / BHK",
        "in_area": "கட்டிடப் பரப்பு (சதுர அடி)", "in_bath": "குளியலறைகள்", "in_balcony": "பால்கனிகள்",
        "result_label": "மதிப்பிடப்பட்ட விற்பனை விலை", "result_ready": "தயார்",
        "result_ready_note": "படிவத்தை நிரப்பி மதிப்பீட்டைப் பெறுங்கள்.",
        "login_title": "BLR HOME IQ-க்கு வரவேற்கிறோம்", "tab_login": "உள்நுழைவு",
        "tab_signup": "பதிவு செய்க", "email": "மின்னஞ்சல் முகவரி", "password": "கடவுச்சொல்",
        "confirm_password": "கடவுச்சொல்லை உறுதிப்படுத்தவும்", "login_btn": "உள்நுழை",
        "signup_btn": "கணக்கை உருவாக்கு", "logout_btn": "வெளியேறு",
        "footer_left": "BLR HOME IQ · பெங்களூரு சொத்து நுண்ணறிவு",
    },
    "తెలుగు": {
        "select_language": "🌐 భాషను ఎంచుకోండి", "country_India": "బెంగళూరు",
        "brand_sub": "బెంగళూరు ప్రాపర్టీ ఇంటెలిజెన్స్", "hero_eyebrow": "AI ధర అంచనా",
        "hero_title": "ప్రతి ఇంటి విలువను<br>సులభంగా తెలుసుకోండి.",
        "hero_text": "ఆస్తి వివరాల ఆధారంగా ధర అంచనాను పొందండి.",
        "nav_home": "హోమ్", "nav_predict": "ఆస్తి ధర అంచనా",
        "nav_history": "నా చరిత్ర", "nav_profile": "ప్రొఫైల్",
        "configure_property": "మీ ఆస్తి వివరాలను నమోదు చేయండి", "your_valuation": "మీ అంచనా",
        "submit_hint": "ధర అంచనా కోసం ఆస్తి వివరాలను నమోదు చేయండి.",
        "estimate_btn": "✦ ఆస్తి ధరను అంచనా వేయండి", "in_location": "ప్రాంతం (బెంగళూరు)",
        "in_area_type": "ఆస్తి రకం", "in_bhk": "బెడ్‌రూమ్‌లు / BHK",
        "in_area": "నిర్మిత విస్తీర్ణం (చ.అ.)", "in_bath": "బాత్‌రూమ్‌లు", "in_balcony": "బాల్కనీలు",
        "result_label": "అంచనా అమ్మకపు ధర", "result_ready": "మీ కోసం సిద్ధంగా ఉంది",
        "result_ready_note": "ఫారమ్ నింపి ధర అంచనాను పొందండి.",
        "login_title": "BLR HOME IQ కు స్వాగతం", "tab_login": "లాగిన్",
        "tab_signup": "సైన్ అప్", "email": "ఇమెయిల్ చిరునామా", "password": "పాస్‌వర్డ్",
        "confirm_password": "పాస్‌వర్డ్‌ను నిర్ధారించండి", "login_btn": "లాగిన్",
        "signup_btn": "ఖాతా సృష్టించండి", "logout_btn": "లాగ్ అవుట్",
        "footer_left": "BLR HOME IQ · బెంగళూరు ప్రాపర్టీ ఇంటెలిజెన్స్",
    },
    "ಕನ್ನಡ": {
        "select_language": "🌐 ಭಾಷೆಯನ್ನು ಆಯ್ಕೆಮಾಡಿ", "country_India": "ಬೆಂಗಳೂರು",
        "brand_sub": "ಬೆಂಗಳೂರು ಆಸ್ತಿ ಮಾಹಿತಿ", "hero_eyebrow": "AI ಬೆಲೆ ಅಂದಾಜು",
        "hero_title": "ಪ್ರತಿ ಮನೆಯ ಮೌಲ್ಯವನ್ನು<br>ಸುಲಭವಾಗಿ ತಿಳಿಯಿರಿ.",
        "hero_text": "ಆಸ್ತಿ ವಿವರಗಳ ಆಧಾರದ ಮೇಲೆ ಬೆಲೆ ಅಂದಾಜು ಪಡೆಯಿರಿ.",
        "nav_home": "ಮುಖಪುಟ", "nav_predict": "ಆಸ್ತಿ ಬೆಲೆ ಅಂದಾಜು",
        "nav_history": "ನನ್ನ ಇತಿಹಾಸ", "nav_profile": "ಪ್ರೊಫೈಲ್",
        "configure_property": "ನಿಮ್ಮ ಆಸ್ತಿ ವಿವರಗಳನ್ನು ನಮೂದಿಸಿ", "your_valuation": "ನಿಮ್ಮ ಅಂದಾಜು",
        "submit_hint": "ಬೆಲೆ ಅಂದಾಜು ನೋಡಲು ಆಸ್ತಿ ವಿವರಗಳನ್ನು ನಮೂದಿಸಿ.",
        "estimate_btn": "✦ ಆಸ್ತಿ ಬೆಲೆಯನ್ನು ಅಂದಾಜಿಸಿ", "in_location": "ಸ್ಥಳ (ಬೆಂಗಳೂರು)",
        "in_area_type": "ಆಸ್ತಿ ಪ್ರಕಾರ", "in_bhk": "ಮಲಗುವ ಕೋಣೆಗಳು / BHK",
        "in_area": "ನಿರ್ಮಿತ ವಿಸ್ತೀರ್ಣ (ಚ.ಅಡಿ)", "in_bath": "ಬಾತ್‌ರೂಮ್‌ಗಳು", "in_balcony": "ಬಾಲ್ಕನಿಗಳು",
        "result_label": "ಅಂದಾಜು ಮಾರಾಟ ಬೆಲೆ", "result_ready": "ನಿಮಗಾಗಿ ಸಿದ್ಧವಾಗಿದೆ",
        "result_ready_note": "ಫಾರ್ಮ್ ಭರ್ತಿ ಮಾಡಿ ಬೆಲೆ ಅಂದಾಜು ಪಡೆಯಿರಿ.",
        "login_title": "BLR HOME IQ ಗೆ ಸ್ವಾಗತ", "tab_login": "ಲಾಗಿನ್",
        "tab_signup": "ಸೈನ್ ಅಪ್", "email": "ಇಮೇಲ್ ವಿಳಾಸ", "password": "ಪಾಸ್‌ವರ್ಡ್",
        "confirm_password": "ಪಾಸ್‌ವರ್ಡ್ ದೃಢೀಕರಿಸಿ", "login_btn": "ಲಾಗಿನ್",
        "signup_btn": "ಖಾತೆ ರಚಿಸಿ", "logout_btn": "ಲಾಗ್ ಔಟ್",
        "footer_left": "BLR HOME IQ · ಬೆಂಗಳೂರು ಆಸ್ತಿ ಮಾಹಿತಿ",
    },
    "മലയാളം": {
        "select_language": "🌐 ഭാഷ തിരഞ്ഞെടുക്കുക", "country_India": "ബെംഗളൂരു",
        "brand_sub": "ബെംഗളൂരു പ്രോപ്പർട്ടി ഇന്റലിജൻസ്", "hero_eyebrow": "AI വില നിർണ്ണയം",
        "hero_title": "ഓരോ വീടിന്റെയും<br>മൂല്യം അറിയൂ.",
        "hero_text": "പ്രോപ്പർട്ടി വിവരങ്ങളുടെ അടിസ്ഥാനത്തിൽ വില കണക്കാക്കുക.",
        "nav_home": "ഹോം", "nav_predict": "പ്രോപ്പർട്ടി വില",
        "nav_history": "എന്റെ ചരിത്രം", "nav_profile": "പ്രൊഫൈൽ",
        "configure_property": "പ്രോപ്പർട്ടി വിവരങ്ങൾ നൽകുക", "your_valuation": "നിങ്ങളുടെ വിലയിരുത്തൽ",
        "submit_hint": "വില അറിയാൻ പ്രോപ്പർട്ടി വിവരങ്ങൾ നൽകുക.",
        "estimate_btn": "✦ പ്രോപ്പർട്ടി വില കണക്കാക്കുക", "in_location": "സ്ഥലം (ബെംഗളൂരു)",
        "in_area_type": "പ്രോപ്പർട്ടി തരം", "in_bhk": "കിടപ്പുമുറികൾ / BHK",
        "in_area": "നിർമിത വിസ്തീർണം (ചതുരശ്ര അടി)", "in_bath": "കുളിമുറികൾ", "in_balcony": "ബാൽക്കണികൾ",
        "result_label": "കണക്കാക്കിയ വിൽപ്പന വില", "result_ready": "തയ്യാറാണ്",
        "result_ready_note": "ഫോം പൂരിപ്പിച്ച് വില കണക്കാക്കുക.",
        "login_title": "BLR HOME IQ-ലേക്ക് സ്വാഗതം", "tab_login": "ലോഗിൻ",
        "tab_signup": "സൈൻ അപ്പ്", "email": "ഇമെയിൽ വിലാസം", "password": "പാസ്‌വേഡ്",
        "confirm_password": "പാസ്‌വേഡ് സ്ഥിരീകരിക്കുക", "login_btn": "ലോഗിൻ",
        "signup_btn": "അക്കൗണ്ട് സൃഷ്ടിക്കുക", "logout_btn": "ലോഗ് ഔട്ട്",
        "footer_left": "BLR HOME IQ · ബെംഗളൂരു പ്രോപ്പർട്ടി ഇന്റലിജൻസ്",
    },
    "বাংলা": {
        "select_language": "🌐 ভাষা নির্বাচন করুন", "country_India": "বেঙ্গালুরু",
        "brand_sub": "বেঙ্গালুরু প্রপার্টি ইন্টেলিজেন্স", "hero_eyebrow": "AI মূল্যায়ন",
        "hero_title": "প্রতিটি বাড়ির<br>মূল্য সহজে জানুন।",
        "hero_text": "বাড়ির বিবরণের ভিত্তিতে মূল্য অনুমান পান।",
        "nav_home": "হোম", "nav_predict": "প্রপার্টির মূল্য",
        "nav_history": "আমার ইতিহাস", "nav_profile": "প্রোফাইল",
        "configure_property": "আপনার প্রপার্টির বিবরণ দিন", "your_valuation": "আপনার মূল্যায়ন",
        "submit_hint": "মূল্য অনুমান দেখতে প্রপার্টির বিবরণ দিন।", "estimate_btn": "✦ প্রপার্টির মূল্য অনুমান করুন",
        "in_location": "অবস্থান (বেঙ্গালুরু)", "in_area_type": "প্রপার্টির ধরন", "in_bhk": "শোবার ঘর / BHK",
        "in_area": "নির্মিত এলাকা (বর্গফুট)", "in_bath": "বাথরুম", "in_balcony": "বারান্দা",
        "result_label": "আনুমানিক বিক্রয় মূল্য", "result_ready": "আপনার জন্য প্রস্তুত",
        "result_ready_note": "ফর্ম পূরণ করে মূল্য অনুমান করুন।", "login_title": "BLR HOME IQ-তে স্বাগতম",
        "tab_login": "লগ ইন", "tab_signup": "সাইন আপ", "email": "ইমেল ঠিকানা", "password": "পাসওয়ার্ড",
        "confirm_password": "পাসওয়ার্ড নিশ্চিত করুন", "login_btn": "লগ ইন", "signup_btn": "অ্যাকাউন্ট তৈরি করুন",
        "logout_btn": "লগ আউট", "footer_left": "BLR HOME IQ · বেঙ্গালুরু প্রপার্টি ইন্টেলিজেন্স",
    },
    "मराठी": {
        "select_language": "🌐 भाषा निवडा", "country_India": "बेंगळुरू",
        "brand_sub": "बेंगळुरू प्रॉपर्टी इंटेलिजन्स", "hero_eyebrow": "AI मूल्यांकन",
        "hero_title": "प्रत्येक घराची किंमत<br>सहज जाणून घ्या.",
        "hero_text": "मालमत्तेच्या तपशीलांवर आधारित किंमतीचा अंदाज मिळवा.",
        "nav_home": "मुख्यपृष्ठ", "nav_predict": "मालमत्तेची किंमत",
        "nav_history": "माझा इतिहास", "nav_profile": "प्रोफाइल",
        "configure_property": "तुमच्या मालमत्तेचे तपशील भरा", "your_valuation": "तुमचा अंदाज",
        "submit_hint": "किंमत पाहण्यासाठी मालमत्तेचे तपशील भरा.", "estimate_btn": "✦ मालमत्तेची किंमत अंदाजित करा",
        "in_location": "ठिकाण (बेंगळुरू)", "in_area_type": "मालमत्तेचा प्रकार", "in_bhk": "बेडरूम / BHK",
        "in_area": "बांधलेले क्षेत्रफळ (चौ. फूट)", "in_bath": "बाथरूम", "in_balcony": "बाल्कनी",
        "result_label": "अंदाजित विक्री किंमत", "result_ready": "तुमच्यासाठी तयार",
        "result_ready_note": "फॉर्म भरा आणि किंमतीचा अंदाज मिळवा.", "login_title": "BLR HOME IQ मध्ये स्वागत आहे",
        "tab_login": "लॉग इन", "tab_signup": "साइन अप", "email": "ईमेल पत्ता", "password": "पासवर्ड",
        "confirm_password": "पासवर्डची पुष्टी करा", "login_btn": "लॉग इन", "signup_btn": "खाते तयार करा",
        "logout_btn": "लॉग आउट", "footer_left": "BLR HOME IQ · बेंगळुरू प्रॉपर्टी इंटेलिजन्स",
    },
    "ગુજરાતી": {
        "select_language": "🌐 ભાષા પસંદ કરો", "country_India": "બેંગલુરુ",
        "brand_sub": "બેંગલુરુ પ્રોપર્ટી ઇન્ટેલિજન્સ", "hero_eyebrow": "AI કિંમત અંદાજ",
        "hero_title": "દરેક ઘરની કિંમત<br>સરળતાથી જાણો.",
        "hero_text": "મિલકતની વિગતોના આધારે કિંમતનો અંદાજ મેળવો.",
        "nav_home": "હોમ", "nav_predict": "મિલકતની કિંમત",
        "nav_history": "મારો ઇતિહાસ", "nav_profile": "પ્રોફાઇલ",
        "configure_property": "તમારી મિલકતની વિગતો ભરો", "your_valuation": "તમારો અંદાજ",
        "submit_hint": "કિંમત જાણવા મિલકતની વિગતો ભરો.", "estimate_btn": "✦ મિલકતની કિંમતનો અંદાજ મેળવો",
        "in_location": "સ્થાન (બેંગલુરુ)", "in_area_type": "મિલકતનો પ્રકાર", "in_bhk": "બેડરૂમ / BHK",
        "in_area": "બાંધકામ વિસ્તાર (ચોરસ ફૂટ)", "in_bath": "બાથરૂમ", "in_balcony": "બાલ્કની",
        "result_label": "અંદાજિત વેચાણ કિંમત", "result_ready": "તમારા માટે તૈયાર",
        "result_ready_note": "ફોર્મ ભરો અને કિંમતનો અંદાજ મેળવો.", "login_title": "BLR HOME IQ માં સ્વાગત છે",
        "tab_login": "લૉગ ઇન", "tab_signup": "સાઇન અપ", "email": "ઈમેલ સરનામું", "password": "પાસવર્ડ",
        "confirm_password": "પાસવર્ડની પુષ્ટિ કરો", "login_btn": "લૉગ ઇન", "signup_btn": "ખાતું બનાવો",
        "logout_btn": "લૉગ આઉટ", "footer_left": "BLR HOME IQ · બેંગલુરુ પ્રોપર્ટી ઇન્ટેલિજન્સ",
    },
    "اردو": {
        "select_language": "🌐 زبان منتخب کریں", "country_India": "بنگلورو",
        "brand_sub": "بنگلورو پراپرٹی انٹیلیجنس", "hero_eyebrow": "AI قیمت کا اندازہ",
        "hero_title": "ہر گھر کی قیمت<br>آسانی سے جانیں۔",
        "hero_text": "پراپرٹی کی تفصیلات کی بنیاد پر قیمت کا اندازہ حاصل کریں۔",
        "nav_home": "ہوم", "nav_predict": "پراپرٹی کی قیمت",
        "nav_history": "میری تاریخ", "nav_profile": "پروفائل",
        "configure_property": "اپنی پراپرٹی کی تفصیلات درج کریں", "your_valuation": "آپ کا تخمینہ",
        "submit_hint": "قیمت دیکھنے کے لیے پراپرٹی کی تفصیلات درج کریں۔", "estimate_btn": "✦ پراپرٹی کی قیمت کا اندازہ لگائیں",
        "in_location": "مقام (بنگلورو)", "in_area_type": "پراپرٹی کی قسم", "in_bhk": "بیڈ روم / BHK",
        "in_area": "تعمیر شدہ رقبہ (مربع فٹ)", "in_bath": "باتھ روم", "in_balcony": "بالکونیاں",
        "result_label": "تخمینی فروخت قیمت", "result_ready": "آپ کے لیے تیار",
        "result_ready_note": "فارم مکمل کریں اور قیمت کا اندازہ حاصل کریں۔", "login_title": "BLR HOME IQ میں خوش آمدید",
        "tab_login": "لاگ ان", "tab_signup": "سائن اپ", "email": "ای میل ایڈریس", "password": "پاس ورڈ",
        "confirm_password": "پاس ورڈ کی تصدیق کریں", "login_btn": "لاگ ان", "signup_btn": "اکاؤنٹ بنائیں",
        "logout_btn": "لاگ آؤٹ", "footer_left": "BLR HOME IQ · بنگلورو پراپرٹی انٹیلیجنس",
    },

}

AREA_TYPE_NAMES_HI = {
    "super built-up area": "सुपर बिल्ट-अप क्षेत्र",
    "plot area": "प्लॉट क्षेत्र",
    "built-up area": "बिल्ट-अप क्षेत्र",
    "carpet area": "कार्पेट क्षेत्र",
}

# Which saved property value goes to which table column
HISTORY_FIELDS = {
    "India": {
        "location": "col_location",
        "area_type": "col_area_type",
        "bhk": "col_bhk",
        "total_sqft": "col_area",
        "bath": "col_bath",
        "balcony": "col_balcony",
    },
}


def current_language():
    lang = st.session_state.get("language", "English")
    return lang if lang in TEXT else "English"


def t(key, **kwargs):
    """Text in the selected language (falls back to English)."""
    text = TEXT[current_language()].get(key, TEXT["English"].get(key, key))
    return text.format(**kwargs) if kwargs else text


def area_type_label(value):
    if current_language() == "English":
        return str(value)
    return AREA_TYPE_NAMES_HI.get(" ".join(str(value).split()).lower(), str(value))


def format_inr(price_in_lakh):
    """Below 100 lakh -> lakh, 100 lakh and above -> crore (in the selected language)."""
    price_in_lakh = float(price_in_lakh)
    if price_in_lakh >= 100:
        return f"₹{price_in_lakh / 100:,.2f} {t('unit_crore')}"
    return f"₹{price_in_lakh:,.2f} {t('unit_lakh')}"


def localize_property_details(details):
    """{readable column name: value} for one saved prediction, in the selected language."""
    result = {}
    for key, label_key in HISTORY_FIELDS["India"].items():
        value = details.get(key, "")
        if value != "":
            if key == "area_type":
                value = area_type_label(value)
        result[t(label_key)] = value
    return result


# ==================== SMALL HTML HELPERS ====================
def esc(value):
    """Escape a dynamic value before it is placed inside HTML."""
    return html_lib.escape(str(value))


def render_html(markup):
    """Render an HTML snippet. Lines are stripped and joined so Markdown never turns
    indented HTML into a code block."""
    flat = "".join(line.strip() for line in markup.splitlines())
    st.markdown(flat, unsafe_allow_html=True)


def go_to_page(page):
    """Button callback: open one dashboard page."""
    st.session_state["page"] = page


def render_divider():
    """Subtle warm divider. Replaces the old st.markdown('---') line."""
    render_html('<div class="soft-divider"></div>')


def render_page_header(title, sub):
    render_html(
        f'<div class="page-head"><div class="section-title">{title}</div>'
        f'<div class="section-sub">{sub}</div></div>'
    )


# ==================== LOGIN / SIGN UP (Supabase Auth) ====================
def new_supabase_client():
    """A fresh client for every action, so one user's login is never shared with another."""
    return create_client(st.secrets["SUPABASE_URL"], st.secrets["SUPABASE_KEY"])


def friendly_auth_error(exc):
    message = str(exc)
    low = message.lower()
    if isinstance(exc, KeyError) or "secrets" in low or "supabase_url" in low or "supabase_key" in low:
        return t("err_secrets")
    if "invalid login credentials" in low:
        return t("err_invalid_login")
    if "email not confirmed" in low:
        return t("err_email_not_confirmed")
    if "already registered" in low or "already been registered" in low or "user_already_exists" in low:
        return t("err_email_exists")
    if "rate limit" in low or "over_email_send_rate_limit" in low or "too many" in low:
        return t("err_rate_limit")
    if "password" in low and ("weak" in low or "at least" in low or "short" in low):
        return t("err_pw_short")
    if "email" in low and "invalid" in low:
        return t("err_email_invalid")
    if any(w in low for w in ("connect", "timed out", "network", "getaddrinfo", "name or service")):
        return t("err_network")
    return t("err_generic", error=message)


def run_auth_action(action, *args):
    """Run a login/sign-up call. Returns (result, error_text)."""
    try:
        return action(*args), None
    except Exception as exc:
        return None, friendly_auth_error(exc)


def do_login(email, password):
    res = new_supabase_client().auth.sign_in_with_password({"email": email, "password": password})
    return res.user, res.session


def do_signup(email, password):
    res = new_supabase_client().auth.sign_up({"email": email, "password": password})
    return res.user, res.session


def set_logged_in(user, session):
    st.session_state["user"] = {
        "id": str(user.id),
        "email": user.email,
        "access_token": session.access_token,
        "refresh_token": session.refresh_token,
    }
    st.session_state["page"] = "home"
    st.session_state.pop("history_cache", None)


def clear_prediction_state():
    for key in ("prediction", "prediction_country", "last_house_data", "_last_country",
                "history_cache", "history_warning", "input_warning"):
        st.session_state.pop(key, None)


def show_login_screen():
    """Login / sign-up page. Nothing else in the app is shown until the user logs in."""
    _, middle, _ = st.columns([1, 1.6, 1])
    with middle:
        st.selectbox(t("select_language"), list(TEXT.keys()), key="language")
        render_page_header(t("login_title"), t("login_sub"))
        if not SUPABASE_AVAILABLE:
            st.error(t("err_no_supabase"))
            return
        if st.session_state.pop("session_expired", False):
            st.warning(t("err_session_expired"))

        login_tab, signup_tab = st.tabs([t("tab_login"), t("tab_signup")])
        logged_in_user = None
        signup_notice = None

        with login_tab:
            with st.form("login_form", border=False):
                login_email = st.text_input(t("email"), key="login_email")
                login_password = st.text_input(t("password"), type="password", key="login_password")
                login_clicked = st.form_submit_button(t("login_btn"), width="stretch")
            if login_clicked:
                email = login_email.strip()
                if "@" not in email or "." not in email:
                    st.error(t("err_email_invalid"))
                elif not login_password:
                    st.error(t("err_pw_short"))
                else:
                    with st.spinner("..."):
                        result, error = run_auth_action(do_login, email, login_password)
                    if error:
                        st.error(error)
                    elif result is None or result[0] is None or result[1] is None:
                        st.error(t("err_invalid_login"))
                    else:
                        logged_in_user = result

        with signup_tab:
            with st.form("signup_form", border=False):
                signup_email = st.text_input(t("email"), key="signup_email")
                signup_password = st.text_input(t("password"), type="password", key="signup_password")
                signup_password2 = st.text_input(t("confirm_password"), type="password", key="signup_password2")
                signup_clicked = st.form_submit_button(t("signup_btn"), width="stretch")
            if signup_clicked:
                email = signup_email.strip()
                if "@" not in email or "." not in email:
                    st.error(t("err_email_invalid"))
                elif len(signup_password) < 8:
                    st.error(t("err_pw_short"))
                elif signup_password != signup_password2:
                    st.error(t("err_pw_mismatch"))
                else:
                    with st.spinner("..."):
                        result, error = run_auth_action(do_signup, email, signup_password)
                    if error:
                        st.error(error)
                    else:
                        user, session = result
                        if user is None:
                            st.error(t("err_generic", error="no user returned"))
                        elif getattr(user, "identities", None) == []:
                            # Supabase hides "already registered" this way when confirmation is on
                            st.error(t("err_email_exists"))
                        elif session is not None:
                            logged_in_user = (user, session)
                        else:
                            signup_notice = t("signup_check_email")

        if signup_notice:
            st.success(signup_notice)

    if logged_in_user is not None:
        set_logged_in(*logged_in_user)
        st.rerun()


# ==================== HISTORY (saved in Supabase, one user at a time) ====================
class SessionExpired(Exception):
    pass


def looks_like_network_error(message):
    low = str(message).lower()
    return any(w in low for w in ("connect", "timed out", "network", "getaddrinfo", "name or service"))


def token_expires_soon(token, margin_seconds=60):
    """Read the expiry time inside the login token (no network call)."""
    try:
        payload = token.split(".")[1]
        payload += "=" * (-len(payload) % 4)
        expiry = json.loads(base64.urlsafe_b64decode(payload))["exp"]
        return expiry - time.time() < margin_seconds
    except Exception:
        return True


def get_user_client():
    """A Supabase client that acts as the logged-in user (so the database only shows THEIR rows)."""
    user = st.session_state["user"]
    client = new_supabase_client()
    if token_expires_soon(user["access_token"]):
        try:
            refreshed = client.auth.refresh_session(user["refresh_token"])
            user["access_token"] = refreshed.session.access_token
            user["refresh_token"] = refreshed.session.refresh_token
        except Exception as exc:
            if looks_like_network_error(exc):
                raise
            raise SessionExpired(str(exc))
    client.postgrest.auth(user["access_token"])
    return client


def history_error_text(exc):
    low = str(exc).lower()
    if isinstance(exc, SessionExpired) or "jwt" in low or "expired" in low:
        return None  # means: session expired
    if looks_like_network_error(low):
        return t("err_network")
    if "does not exist" in low or "pgrst205" in low:
        return "table prediction_history not found in Supabase"
    return str(exc)


def force_logout_expired():
    st.session_state.pop("user", None)
    clear_prediction_state()
    st.session_state["session_expired"] = True
    st.rerun()


def save_prediction_history(property_data, price):
    """Save one prediction in the logged-in user's history. Returns an error text or None."""
    details = json.loads(json.dumps(property_data.iloc[0].to_dict(), default=str))
    try:
        get_user_client().table("prediction_history").insert({
            "country": "India",
            "estimated_price": round(float(price), 2),
            "currency": "INR lakh",
            "property_details": details,
        }).execute()
    except Exception as exc:
        st.session_state.pop("history_cache", None)
        text = history_error_text(exc)
        if text is None:
            force_logout_expired()
        return text
    st.session_state.pop("history_cache", None)
    return None


def load_prediction_history():
    """Return India history and an error text. Oldest first; only this user's rows."""
    cache = st.session_state.setdefault("history_cache", {})
    if "India" in cache:
        return cache["India"], None
    try:
        response = (
            get_user_client().table("prediction_history")
            .select("created_at,country,estimated_price,currency,property_details")
            .eq("country", "India")
            .order("created_at", desc=True)
            .limit(1000)
            .execute()
        )
    except Exception as exc:
        text = history_error_text(exc)
        if text is None:
            force_logout_expired()
        return pd.DataFrame(columns=HISTORY_COLUMNS), text

    rows = []
    for r in reversed(response.data or []):
        try:
            shown_time = datetime.fromisoformat(r["created_at"].replace("Z", "+00:00")).astimezone()
            shown_time = shown_time.strftime("%Y-%m-%d %H:%M:%S")
        except (ValueError, AttributeError):
            shown_time = str(r.get("created_at", ""))
        rows.append({
            "Date and time": shown_time,
            "Country": r.get("country", "India"),
            "Estimated price": r.get("estimated_price"),
            "Currency": r.get("currency", ""),
            "Property details": json.dumps(r.get("property_details") or {}, ensure_ascii=False),
        })
    df = pd.DataFrame(rows, columns=HISTORY_COLUMNS)
    cache["India"] = df
    return df, None


def clear_prediction_history():
    """Delete only this user's India history. Returns an error text or None."""
    try:
        get_user_client().table("prediction_history").delete().eq("country", "India").execute()
    except Exception as exc:
        text = history_error_text(exc)
        if text is None:
            force_logout_expired()
        return text
    st.session_state.pop("history_cache", None)
    return None


def format_history_for_display(df):
    """Readable table in the selected language. Only changes what is shown;
    the saved CSV is not modified."""
    rows = []
    for _, r in df.iterrows():
        try:
            details = json.loads(r["Property details"])
        except (TypeError, ValueError):
            details = {}
        try:
            price = float(r["Estimated price"])
            price_text = format_inr(price)
        except (TypeError, ValueError):
            price_text = str(r["Estimated price"])
        row = {t("col_datetime"): r["Date and time"], t("col_price"): price_text}
        row.update(localize_property_details(details))
        rows.append(row)
    return pd.DataFrame(rows)


def history_to_excel_bytes(df, sheet_name):
    """Build a neat Excel table: coloured header, auto column width, filter, frozen header."""
    buffer = BytesIO()
    with pd.ExcelWriter(buffer, engine="openpyxl") as writer:
        df.to_excel(writer, index=False, sheet_name=sheet_name)
        ws = writer.sheets[sheet_name]

        # Brand colours: forest green header, warm cream zebra rows.
        header_fill = PatternFill("solid", fgColor="0B3D2E")
        header_font = Font(bold=True, color="FFFFFF", size=11)
        thin = Side(style="thin", color="C9D9CE")
        border = Border(left=thin, right=thin, top=thin, bottom=thin)
        zebra_fill = PatternFill("solid", fgColor="F7F1E1")

        for cell in ws[1]:
            cell.fill = header_fill
            cell.font = header_font
            cell.border = border
            cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
        ws.row_dimensions[1].height = 28

        for row_idx, row in enumerate(ws.iter_rows(min_row=2), start=2):
            for cell in row:
                cell.border = border
                cell.alignment = Alignment(horizontal="center", vertical="center")
                if row_idx % 2 == 1:
                    cell.fill = zebra_fill

        for col_idx, column in enumerate(df.columns, start=1):
            longest = max(
                [len(str(column))] + [len(str(v)) for v in df[column].tolist()]
            )
            ws.column_dimensions[get_column_letter(col_idx)].width = min(max(longest + 4, 12), 40)

        ws.freeze_panes = "A2"
        ws.auto_filter.ref = ws.dimensions
    return buffer.getvalue()


@st.cache_resource
def load_model(path):
    # Only load model files that you created or trust: pickle/joblib can run code.
    return joblib.load(path)


def model_display_name(model):
    estimator = model.steps[-1][1] if hasattr(model, "steps") else model
    name = type(estimator).__name__
    return {"RandomForestRegressor": "Random Forest",
            "ExtraTreesRegressor": "Extra Trees"}.get(name, name.replace("Regressor", ""))


@st.cache_data
def load_india_locations(path):
    data = pd.read_csv(path)
    return sorted(data["location"].dropna().astype(str).unique().tolist())


@st.cache_data
def load_area_types(path):
    # Dataset se exact property type labels lena (cached, so the CSV is not re-read on every click)
    data = pd.read_csv(path)
    return sorted(data["area_type"].dropna().astype(str).unique().tolist())


# -------------------- STYLING --------------------
# One stylesheet for the whole app. Colours: deep forest green, warm cream, saffron.
APP_CSS = """
@import url('https://fonts.googleapis.com/css2?family=DM+Sans:wght@400;500;600;700&family=Manrope:wght@500;600;700;800&display=swap');
:root {
 --forest:#0b3d2e; --forest-2:#0f5132; --forest-3:#1a6b4a;
 --cream:#fbf6ea; --cream-2:#f4ecd8; --paper:#fffdf8;
 --saffron:#f28c1b; --saffron-2:#e0700e;
 --ink:#143026; --muted:#5f7468;
 --line:rgba(15,81,50,.15); --shadow:0 10px 30px rgba(11,61,46,.08);
}
.stApp, .stApp p, .stApp label, .stApp li, .stApp input, .stApp textarea, .stApp button,
.stApp h1, .stApp h2, .stApp h3, .stApp h4 { font-family:'DM Sans',system-ui,-apple-system,'Segoe UI',sans-serif; }
.stApp {
 color:var(--ink); color-scheme:light;
 background:radial-gradient(900px 420px at 88% -8%,rgba(242,140,27,.13),transparent 60%),
  linear-gradient(160deg,#fbf6ea 0%,#f3f6ec 58%,#fdf0dc 100%);
}
.stApp :where(p, li, label) { color:var(--ink); }
[data-testid="stDecoration"] { background:linear-gradient(90deg,#0b3d2e,#f28c1b); height:3px; }
[data-testid="stHeader"] { background:transparent; }
.block-container { max-width:1240px; padding-top:1.4rem; padding-bottom:3rem; }
/* The old navy line came from a Markdown horizontal rule. It is replaced by .soft-divider;
   any other horizontal rule is restyled to a soft green hairline. */
.stApp hr { border:0; height:1px; background:var(--line); margin:1.4rem 0; }
.soft-divider { height:1px; margin:30px 0 24px; background:linear-gradient(90deg,transparent,rgba(15,81,50,.22),transparent); }

/* ---------- sidebar ---------- */
section[data-testid="stSidebar"] { background:linear-gradient(185deg,#0b3d2e 0%,#093024 100%); border-right:0; }
section[data-testid="stSidebar"] :is(p,span,label,div) { color:#fbf6ea; }
section[data-testid="stSidebar"] > div { padding-top:.4rem; }
.side-brand { display:flex; align-items:center; gap:12px; padding:14px 6px 8px; }
.side-mark {
 width:46px; height:46px; flex:0 0 46px; border-radius:14px; display:flex; align-items:center; justify-content:center;
 background:linear-gradient(135deg,#1a6b4a,#f28c1b); color:#fffdf8; font-size:24px; font-weight:800;
 box-shadow:0 8px 24px rgba(0,0,0,.25);
}
.side-name { font-family:'Manrope',sans-serif; font-size:21px; font-weight:800; letter-spacing:-.4px; line-height:1.1; color:#fbf6ea; }
.side-name .iq { color:#f9a43a; }
section[data-testid="stSidebar"] .side-tag { font-size:12px; color:rgba(251,246,234,.72); margin-top:3px; }
.side-skyline { opacity:.35; margin:6px 0 4px; }
section[data-testid="stSidebar"] .stButton button {
 justify-content:flex-start; text-align:left; background:transparent; border:0; box-shadow:none;
 color:#fbf6ea; border-radius:12px; min-height:46px; padding:0 14px; font-weight:600; gap:.6rem;
}
section[data-testid="stSidebar"] .stButton button:hover { background:rgba(255,255,255,.09); color:#ffffff; }
section[data-testid="stSidebar"] .stButton button[kind="primary"],
section[data-testid="stSidebar"] .stButton button[data-testid="stBaseButton-primary"] {
 background:linear-gradient(90deg,#1c7350,#15583d); color:#ffffff; box-shadow:inset 0 0 0 1px rgba(255,255,255,.14);
}
section[data-testid="stSidebar"] .st-key-logout_btn button {
 border:1px solid rgba(249,164,58,.65); color:#f9b25a; justify-content:center; min-height:42px;
}
section[data-testid="stSidebar"] .st-key-logout_btn button:hover { background:rgba(249,164,58,.14); color:#ffd08a; }
section[data-testid="stSidebar"] div[data-baseweb="select"] > div,
section[data-testid="stSidebar"] [data-testid="stSelectbox"] [role="group"] {
 background:#12503b !important; border:1px solid rgba(255,255,255,.24) !important; border-radius:12px;
}
section[data-testid="stSidebar"] div[data-baseweb="select"] :is(div,span,input),
section[data-testid="stSidebar"] input {
 color:#fbf6ea !important; -webkit-text-fill-color:#fbf6ea !important; caret-color:#f9a43a !important; background:#12503b !important;
}
section[data-testid="stSidebar"] input::placeholder { color:rgba(251,246,234,.6) !important; -webkit-text-fill-color:rgba(251,246,234,.6) !important; }
section[data-testid="stSidebar"] [data-testid="stSelectbox"] svg, section[data-testid="stSidebar"] div[data-baseweb="select"] svg { fill:#fbf6ea !important; color:#fbf6ea !important; }
.side-user {
 display:flex; align-items:center; gap:12px; padding:12px; margin:10px 0 8px;
 background:rgba(255,255,255,.07); border:1px solid rgba(255,255,255,.12); border-radius:14px;
}
.side-avatar {
 width:38px; height:38px; flex:0 0 38px; border-radius:50%; display:flex; align-items:center; justify-content:center;
 background:#fbf6ea; color:#0b3d2e !important; font-weight:800; font-size:16px;
}
section[data-testid="stSidebar"] .side-user-name { font-size:12px; color:rgba(251,246,234,.7); }
.side-user-mail { font-size:13px; font-weight:600; color:#fbf6ea; overflow-wrap:anywhere; }

/* ---------- hero ---------- */
.hero {
 position:relative; overflow:hidden; border:1px solid rgba(15,81,50,.16); border-radius:28px;
 min-height:340px; padding:38px 40px; margin-bottom:26px; box-shadow:var(--shadow);
 background:linear-gradient(100deg,#fffaf0 0%,#fbf1dc 52%,#f8dcae 100%);
}
.hero-art { position:absolute; right:0; bottom:0; width:58%; height:100%; pointer-events:none; }
.hero-art svg { width:100%; height:100%; display:block; }
.hero-copy { position:relative; z-index:2; max-width:600px; }
.eyebrow {
 display:inline-block; color:var(--forest-2); font-size:11px; font-weight:700; letter-spacing:2px; text-transform:uppercase;
 border:1px solid rgba(15,81,50,.24); background:rgba(15,81,50,.07); padding:7px 12px; border-radius:99px;
}
.hero-welcome { font-family:'Manrope',sans-serif; font-size:clamp(20px,2.4vw,28px); font-weight:700; color:var(--forest); margin-top:16px; }
.hero-brand {
 font-family:'Manrope',sans-serif; font-size:clamp(38px,6vw,68px); line-height:1.02; font-weight:800; letter-spacing:-2px; color:var(--forest);
}
.hero-brand .home { color:var(--saffron); }
.hero-title { font-family:'Manrope',sans-serif; font-size:clamp(17px,2vw,22px); font-weight:700; color:var(--forest-2); margin:12px 0 8px; line-height:1.3; }
.hero-text { color:#4d665b; font-size:15px; line-height:1.65; max-width:500px; margin:0; }
.hero-tags { display:flex; flex-wrap:wrap; gap:9px; margin-top:20px; }
.hero-tag {
 color:var(--forest); border:1px solid rgba(15,81,50,.18); background:rgba(255,255,255,.75);
 border-radius:10px; padding:8px 12px; font-size:12.5px; font-weight:600;
}

/* ---------- headings ---------- */
.section-title { font-family:'Manrope',sans-serif; font-size:22px; font-weight:800; color:var(--forest); letter-spacing:-.4px; margin-bottom:4px; }
.section-sub { color:var(--muted); font-size:13.5px; margin-bottom:16px; line-height:1.5; }
.page-head { margin:4px 0 14px; }

/* ---------- form card ---------- */
div[data-testid="stForm"] {
 border:1px solid var(--line); border-radius:22px; padding:24px 24px 14px;
 background:rgba(255,253,248,.92); box-shadow:var(--shadow);
}
.form-sec { display:flex; align-items:center; gap:12px; margin:6px 0 12px; }
.form-sec-num {
 width:30px; height:30px; flex:0 0 30px; border-radius:50%; display:flex; align-items:center; justify-content:center;
 background:var(--forest); color:#fff; font-weight:800; font-size:14px;
}
.form-sec-title { font-family:'Manrope',sans-serif; font-weight:800; font-size:16px; color:var(--forest); line-height:1.2; }
.form-sec-sub { font-size:12.5px; color:var(--muted); }
.form-gap { height:10px; }
.form-foot { display:flex; flex-wrap:wrap; justify-content:center; gap:8px 18px; margin:2px 0 10px; font-size:12px; color:var(--muted); }
.form-foot b { color:var(--forest-2); font-weight:700; }

/* ---------- inputs ---------- */
div[data-testid="stNumberInput"] label p, div[data-testid="stSelectbox"] label p, div[data-testid="stTextInput"] label p,
div[data-testid="stWidgetLabel"] p { color:var(--forest-2) !important; font-weight:600 !important; font-size:13.5px !important; }
div[data-baseweb="input"] > div, div[data-baseweb="base-input"], div[data-baseweb="select"] > div {
 background:var(--paper); border:1px solid #cddacf; border-radius:12px; min-height:44px;
}
div[data-baseweb="input"]:focus-within > div, div[data-baseweb="select"] > div:focus-within { border-color:var(--saffron); box-shadow:0 0 0 3px rgba(242,140,27,.18); }
.stApp input, .stApp textarea {
 background:#fffdf8 !important; color:#143026 !important; -webkit-text-fill-color:#143026 !important; caret-color:#e0700e !important;
}
.stApp input::placeholder, .stApp textarea::placeholder { color:#8a9a90 !important; -webkit-text-fill-color:#8a9a90 !important; }
.stApp div[data-baseweb="select"] :is(div,span) { color:var(--ink); }
.stApp :is([data-testid="stSelectbox"],[data-testid="stNumberInput"],[data-testid="stTextInput"]) [role="group"] {
 background:#fffdf8 !important; border:1px solid #cddacf; border-radius:12px; min-height:44px;
}
.stApp :is([data-testid="stSelectbox"],[data-testid="stNumberInput"],[data-testid="stTextInput"]) [role="group"]:focus-within {
 border-color:var(--saffron); box-shadow:0 0 0 3px rgba(242,140,27,.18);
}
.stApp [data-testid="stSelectbox"] svg, .stApp [data-testid="stTextInput"] svg { fill:#143026; color:#143026; }
div[data-testid="stNumberInput"] button { background:#f1f4ea; color:var(--forest); border:0; }
div[data-baseweb="popover"] ul, div[data-baseweb="popover"] li { background:var(--paper); color:var(--ink); }
div[data-baseweb="popover"] li:hover { background:#f6ead0; }
[role="listbox"] { background:var(--paper); }
[role="listbox"] [role="option"], [role="listbox"] [role="option"] * { color:var(--ink) !important; }
[role="listbox"] [role="option"]:hover, [role="listbox"] [role="option"][data-focused="true"] { background:#f6ead0; }

/* ---------- buttons ---------- */
.stButton button, .stDownloadButton button {
 border-radius:12px; min-height:46px; font-weight:700; color:var(--forest); background:var(--paper);
 border:1px solid rgba(15,81,50,.28); box-shadow:none;
}
.stButton button:hover, .stDownloadButton button:hover { border-color:var(--saffron); color:var(--saffron-2); background:#fff7ea; }
div[data-testid="stFormSubmitButton"] button {
 border:0; min-height:52px; font-size:16px; color:#ffffff; font-weight:800; border-radius:14px;
 background:linear-gradient(100deg,#f7a02e,#e0700e); box-shadow:0 10px 24px rgba(224,112,14,.28);
}
div[data-testid="stFormSubmitButton"] button:hover { filter:brightness(1.06); color:#ffffff; background:linear-gradient(100deg,#f7a02e,#e0700e); border:0; }
div[data-testid="stFormSubmitButton"] button p { color:#ffffff; }

/* ---------- result card ---------- */
.res-card { border-radius:22px; overflow:hidden; background:var(--paper); border:1px solid var(--line); box-shadow:var(--shadow); }
.res-head { padding:20px 24px; color:#fff; background:radial-gradient(circle at 100% 0%,rgba(255,255,255,.22),transparent 45%),linear-gradient(110deg,#f28c1b,#c9560f); }
.res-head-title { font-family:'Manrope',sans-serif; font-weight:800; font-size:19px; color:#fff; }
.res-head-sub { font-size:13px; color:rgba(255,255,255,.88); margin-top:2px; }
.res-body { padding:22px 24px 24px; }
.result-label { color:var(--muted); text-transform:uppercase; letter-spacing:1.6px; font-size:11px; font-weight:700; text-align:center; }
.result-price {
 font-family:'Manrope',sans-serif; font-size:clamp(32px,4vw,46px); color:var(--forest); font-weight:800;
 letter-spacing:-1.2px; margin:8px 0 6px; overflow-wrap:anywhere; text-align:center;
}
.result-note { display:flex; align-items:center; justify-content:center; gap:8px; background:#e8f3ea; color:var(--forest-2); border-radius:12px; padding:10px 12px; font-size:12.5px; font-weight:600; text-align:center; }
.result-section { font-family:'Manrope',sans-serif; font-weight:800; font-size:13px; color:var(--forest); margin:18px 0 8px; letter-spacing:.2px; }
.chip-row { display:grid; grid-template-columns:repeat(4,1fr); gap:8px; }
.chip { background:#f6f1e2; border-radius:12px; padding:10px 8px; text-align:center; min-width:0; }
.chip span { display:block; font-size:10.5px; color:var(--muted); text-transform:uppercase; letter-spacing:.6px; overflow-wrap:anywhere; }
.chip b { font-family:'Manrope',sans-serif; font-size:19px; color:var(--forest); }
.meta-grid { display:grid; grid-template-columns:1fr 1fr; gap:10px; }
.meta-item { background:#f1f6ee; border-radius:14px; padding:12px 14px; min-width:0; }
.meta-item .k { font-size:11px; color:var(--muted); text-transform:uppercase; letter-spacing:.8px; }
.meta-item .v { font-family:'Manrope',sans-serif; font-weight:800; font-size:15px; color:var(--forest); margin-top:3px; overflow-wrap:anywhere; }
.res-empty-title { font-family:'Manrope',sans-serif; font-size:24px; font-weight:800; color:var(--forest); text-align:center; margin-top:4px; }
.res-empty-note { color:var(--muted); font-size:13.5px; text-align:center; margin:4px 0 16px; }
.steps { list-style:none; margin:0; padding:0; counter-reset:s; }
.steps li { counter-increment:s; display:flex; align-items:center; gap:12px; padding:10px 12px; margin-bottom:8px; background:#f6f1e2; border-radius:12px; font-size:13.5px; color:var(--ink); }
.steps li::before { content:counter(s); width:24px; height:24px; flex:0 0 24px; border-radius:50%; background:var(--saffron); color:#fff; font-weight:800; font-size:12px; display:flex; align-items:center; justify-content:center; }

/* ---------- cards on home / profile / history ---------- */
.qcard { display:flex; gap:14px; align-items:flex-start; background:rgba(255,253,248,.92); border:1px solid var(--line); border-radius:18px; padding:18px; box-shadow:var(--shadow); min-height:104px; }
.qcard-ico { width:44px; height:44px; flex:0 0 44px; border-radius:14px; background:#e6f1e8; display:flex; align-items:center; justify-content:center; font-size:21px; }
.qcard-title { font-family:'Manrope',sans-serif; font-weight:800; font-size:16px; color:var(--forest); }
.qcard-text { font-size:12.5px; color:var(--muted); margin-top:3px; line-height:1.45; }
.panel { background:rgba(255,253,248,.92); border:1px solid var(--line); border-radius:20px; padding:20px 22px; box-shadow:var(--shadow); }
.lang-row { display:flex; flex-wrap:wrap; gap:8px; margin-top:12px; }
.lang-chip { background:#fff; border:1px solid var(--line); border-radius:99px; padding:7px 14px; font-size:13.5px; font-weight:600; color:var(--forest); }
.lang-chip.on { background:var(--forest); color:#fff; border-color:var(--forest); }
.kv { display:flex; justify-content:space-between; gap:16px; padding:10px 0; border-bottom:1px solid var(--line); font-size:14px; }
.kv:last-child { border-bottom:0; }
.kv .k { color:var(--muted); }
.kv .v { color:var(--forest); font-weight:700; text-align:right; overflow-wrap:anywhere; }
.empty-state { text-align:center; padding:34px 20px; background:rgba(255,253,248,.92); border:1px dashed rgba(15,81,50,.3); border-radius:20px; }
.empty-ico { font-size:34px; }
.empty-title { font-family:'Manrope',sans-serif; font-weight:800; font-size:19px; color:var(--forest); margin-top:6px; }
.empty-text { color:var(--muted); font-size:13.5px; margin:6px auto 0; max-width:420px; }
.count-pill { display:inline-block; background:var(--forest); color:#fff; border-radius:99px; padding:6px 14px; font-size:12.5px; font-weight:700; margin-bottom:10px; }

/* ---------- streamlit widgets ---------- */
div[data-testid="stExpander"] { border:1px solid var(--line); border-radius:16px; background:rgba(255,253,248,.9); }
div[data-testid="stExpander"] summary, div[data-testid="stExpander"] summary p { color:var(--forest); font-weight:700; }
div[data-testid="stAlert"] { border-radius:14px; }
div[data-testid="stAlert"] p { color:var(--ink); }
button[data-baseweb="tab"] p { color:var(--muted); font-weight:700; }
button[data-baseweb="tab"][aria-selected="true"] p { color:var(--forest); }
div[data-baseweb="tab-highlight"] { background:var(--saffron) !important; }
div[data-baseweb="tab-border"] { background:var(--line) !important; }
.footer { border-top:1px solid var(--line); margin-top:34px; padding-top:18px; color:var(--muted); font-size:12px; display:flex; justify-content:space-between; gap:12px; flex-wrap:wrap; }

/* ---------- responsive ---------- */
@media (max-width:900px) {
 .hero-art { width:100%; opacity:.28; }
 .hero { padding:28px 22px; min-height:320px; }
}
@media (max-width:640px) {
 .block-container { padding:1rem .8rem 2rem; }
 .hero-brand { letter-spacing:-1px; }
 .chip-row { grid-template-columns:repeat(2,1fr); }
 .meta-grid { grid-template-columns:1fr; }
 div[data-testid="stForm"] { padding:18px 14px 8px; }
 .res-body, .res-head { padding-left:16px; padding-right:16px; }
}
"""


def inject_css():
    # Blank lines are removed so the Markdown parser never ends the <style> block early.
    css = "\n".join(line for line in APP_CSS.splitlines() if line.strip())
    st.markdown(f"<style>{css}</style>", unsafe_allow_html=True)


def hero_art_svg():
    """Original flat illustration: a domed Bengaluru-style government building, towers and trees."""
    windows = "".join(
        f'<rect x="{232 + i * 24}" y="236" width="12" height="30" rx="6"/>' for i in range(11)
    )
    upper_windows = "".join(
        f'<rect x="{298 + i * 22}" y="176" width="10" height="22" rx="5"/>' for i in range(6)
    )
    return (
        '<svg viewBox="0 0 560 330" preserveAspectRatio="xMaxYMax slice" xmlns="http://www.w3.org/2000/svg" aria-hidden="true">'
        '<circle cx="430" cy="130" r="86" fill="#f28c1b" opacity=".22"/>'
        '<g fill="#1a6b4a" opacity=".30">'
        '<rect x="30" y="170" width="36" height="160"/><rect x="74" y="130" width="30" height="200"/>'
        '<rect x="112" y="185" width="42" height="145"/><rect x="162" y="150" width="28" height="180"/></g>'
        '<g fill="#0f5132">'
        '<rect x="214" y="226" width="304" height="104"/><rect x="284" y="166" width="164" height="68"/>'
        '<rect x="336" y="118" width="60" height="56"/><path d="M336 118 Q366 40 396 118 Z"/>'
        '<rect x="363" y="26" width="6" height="34"/>'
        '<path d="M220 226 Q246 184 272 226 Z"/><path d="M460 226 Q486 184 512 226 Z"/>'
        '<path d="M292 166 Q314 130 336 166 Z"/><path d="M396 166 Q418 130 440 166 Z"/></g>'
        f'<g fill="#fbf6ea" opacity=".8">{windows}{upper_windows}</g>'
        '<g fill="#1a6b4a" opacity=".65"><circle cx="540" cy="262" r="46"/><circle cx="196" cy="290" r="34"/>'
        '<circle cx="520" cy="300" r="40"/></g>'
        '<rect x="0" y="318" width="560" height="12" fill="#0b3d2e" opacity=".5"/>'
        '</svg>'
    )


def render_hero():
    render_html(f"""
    <section class="hero">
      <div class="hero-art">{hero_art_svg()}</div>
      <div class="hero-copy">
        <span class="eyebrow">{t('hero_eyebrow')}</span>
        <div class="hero-welcome">{t('hero_welcome')}</div>
        <div class="hero-brand">BLR <span class="home">HOME IQ</span></div>
        <div class="hero-title">{t('hero_title')}</div>
        <p class="hero-text">{t('hero_text')}</p>
        <div class="hero-tags">
          <span class="hero-tag">{t('hero_tag1')}</span>
          <span class="hero-tag">{t('hero_tag2')}</span>
          <span class="hero-tag">{t('hero_tag3')}</span>
        </div>
      </div>
    </section>
    """)


# ==================== SIDEBAR ====================
def render_sidebar():
    user_email = st.session_state["user"]["email"] or ""
    with st.sidebar:
        render_html(f"""
        <div class="side-brand">
          <div class="side-mark">⌂</div>
          <div>
            <div class="side-name">BLR HOME <span class="iq">IQ</span></div>
            <div class="side-tag">{t('nav_tagline')}</div>
          </div>
        </div>
        """)
        current = st.session_state.get("page", "home")
        for page in PAGES:
            st.button(
                t(PAGE_LABEL_KEYS[page]),
                key=f"nav_{page}",
                icon=PAGE_ICONS[page],
                type="primary" if current == page else "secondary",
                width="stretch",
                on_click=go_to_page,
                args=(page,),
            )

        st.selectbox(t("select_language"), list(TEXT.keys()), key="language")

        initial = esc(user_email[:1].upper() or "?")
        render_html(f"""
        <div class="side-user">
          <div class="side-avatar">{initial}</div>
          <div>
            <div class="side-user-name">{t('welcome_user')}</div>
            <div class="side-user-mail">{esc(user_email)}</div>
          </div>
        </div>
        """)
        if st.button(t("logout_btn"), key="logout_btn", icon=":material/logout:", width="stretch"):
            st.session_state.pop("user", None)
            st.session_state.pop("page", None)
            clear_prediction_state()
            st.rerun()


# ==================== PREDICTION WORKSPACE (form + result) ====================
def render_result_panel():
    if "prediction" not in st.session_state:
        render_html(f"""
        <div class="res-card">
          <div class="res-head">
            <div class="res-head-title">{t('result_label')}</div>
            <div class="res-head-sub">{t('submit_hint')}</div>
          </div>
          <div class="res-body">
            <div class="res-empty-title">{t('result_ready')}</div>
            <div class="res-empty-note">{t('result_ready_note')}</div>
            <ol class="steps"><li>{t('step1')}</li><li>{t('step2')}</li><li>{t('step3')}</li></ol>
          </div>
        </div>
        """)
        return

    price = st.session_state["prediction"]
    last = st.session_state.get("last_house_data")
    details = last.iloc[0].to_dict() if last is not None else {}

    summary_html = ""
    rate_html = ""
    if details:
        try:
            summary_html = (
                f'<div class="result-section">{t("summary_title")}</div><div class="chip-row">'
                f'<div class="chip"><span>{t("col_bhk")}</span><b>{esc(int(details.get("bhk")))}</b></div>'
                f'<div class="chip"><span>{t("col_area")}</span><b>{esc(int(details.get("total_sqft")))}</b></div>'
                f'<div class="chip"><span>{t("col_bath")}</span><b>{esc(int(details.get("bath")))}</b></div>'
                f'<div class="chip"><span>{t("col_balcony")}</span><b>{esc(int(details.get("balcony")))}</b></div>'
                f'</div>'
            )
            sqft = float(details.get("total_sqft"))
            if sqft > 0:
                # Simple arithmetic on the model result: lakh -> rupees, divided by the entered area.
                rate_html = (
                    f'<div class="meta-item"><div class="k">{t("stat_rate")}</div>'
                    f'<div class="v">₹{price * 100000 / sqft:,.0f}</div></div>'
                )
        except (TypeError, ValueError):
            summary_html = ""
            rate_html = ""

    location_html = ""
    if details.get("location") not in (None, ""):
        location_html = (
            f'<div class="meta-item"><div class="k">{t("stat_location")}</div>'
            f'<div class="v">{esc(details.get("location"))}</div></div>'
        )

    render_html(f"""
    <div class="res-card">
      <div class="res-head">
        <div class="res-head-title">{t('result_label')}</div>
        <div class="res-head-sub">{t('your_valuation')}</div>
      </div>
      <div class="res-body">
        <div class="result-label">{t('result_label')}</div>
        <div class="result-price">{esc(format_inr(price))}</div>
        <div class="result-note">🛡️ {t('note_india')}</div>
        {summary_html}
        <div class="result-section">{t('stat_model')}</div>
        <div class="meta-grid">
          <div class="meta-item"><div class="k">{t('stat_model')}</div><div class="v">{html_lib.escape(st.session_state.get('model_display_name', t('model_extra_trees')))}</div></div>
          {location_html}
          <div class="meta-item"><div class="k">{t('stat_inputs')}</div><div class="v">{t('inputs_india')}</div></div>
          <div class="meta-item"><div class="k">{t('stat_dataset')}</div><div class="v">{t('dataset_india')}</div></div>
          {rate_html}
        </div>
      </div>
    </div>
    """)


def render_prediction_workspace(india_locations, area_type_options, india_model):
    """The property form (left) and the result dashboard (right)."""
    left, right = st.columns([1.25, 1], gap="large")
    submitted = False

    # Every input has a key so typed values are kept when the language is changed.
    with left:
        with st.form("india_house_form", border=False):
            render_html(f"""
            <div class="section-title">{t('configure_property')}</div>
            <div class="section-sub">{t('india_sub')}</div>
            <div class="form-sec"><div class="form-sec-num">1</div>
              <div><div class="form-sec-title">{t('sec_location')}</div><div class="form-sec-sub">{t('sec_location_sub')}</div></div></div>
            """)
            loc_col, type_col = st.columns(2)
            with loc_col:
                location = st.selectbox(
                    t("in_location"), india_locations, key="in_location", help=t("help_location")
                )
            with type_col:
                area_type = st.selectbox(
                    t("in_area_type"),
                    area_type_options,
                    format_func=area_type_label,
                    key="in_area_type",
                    help=t("help_area_type"),
                )

            render_html(f"""
            <div class="form-gap"></div>
            <div class="form-sec"><div class="form-sec-num">2</div>
              <div><div class="form-sec-title">{t('sec_size')}</div><div class="form-sec-sub">{t('sec_size_sub')}</div></div></div>
            """)
            bhk_col, area_col = st.columns(2)
            with bhk_col:
                bhk = st.number_input(t("in_bhk"), 1, 10, 2, 1, key="in_bhk", help=t("help_bhk"))
            with area_col:
                area = st.number_input(t("in_area"), 300, 10000, 1056, 50, key="in_area", help=t("help_area"))
            bath_col, balcony_col = st.columns(2)
            with bath_col:
                bath = st.number_input(t("in_bath"), 1, 10, 2, 1, key="in_bath", help=t("help_bath"))
            with balcony_col:
                balcony = st.number_input(t("in_balcony"), 0, 5, 1, 1, key="in_balcony", help=t("help_balcony"))

            submitted = st.form_submit_button(
                t("estimate_btn"), type="primary", width="stretch"
            )
            render_html(f"""
            <div class="form-foot">
              <span>{t('stat_model')}: <b>{html_lib.escape(st.session_state.get('model_display_name', t('model_extra_trees')))}</b></span>
              <span>{t('stat_dataset')}: <b>{t('dataset_india')}</b></span>
              <span><b>{t('inputs_india')}</b></span>
            </div>
            """)

        small_area = st.session_state.get("input_warning")
        if small_area:
            st.warning(t("warn_small_area", area=small_area[0], bhk=small_area[1]))

    # -------------------- PREDICTION --------------------
    if submitted:
        try:
            house_data = pd.DataFrame([{
                "location": location,
                "area_type": area_type,
                "bhk": bhk,
                "total_sqft": area,
                "bath": bath,
                "balcony": balcony
            }])

            with st.spinner(t("spinner_predict")):
                prediction_value = float(india_model.predict(house_data)[0])
                if not math.isfinite(prediction_value):
                    raise ValueError("Model returned an invalid price")
                prediction_value = max(prediction_value, 0.0)
                st.session_state["prediction"] = prediction_value
                st.session_state["prediction_country"] = "India"
                st.session_state["last_house_data"] = house_data
                st.session_state["history_warning"] = save_prediction_history(house_data, prediction_value)
            if area / max(bhk, 1) < MIN_SQFT_PER_BHK:
                st.session_state["input_warning"] = (int(area), int(bhk))
            else:
                st.session_state.pop("input_warning", None)
            st.rerun()

        except Exception as exc:
            st.error(t("err_prediction"))
            with st.expander(t("tech_details")):
                st.exception(exc)

    # -------------------- RESULTS PANEL --------------------
    with right:
        render_result_panel()
        save_warning = st.session_state.pop("history_warning", None)
        if save_warning:
            st.warning(t("warn_history_save", error=save_warning))

    if "last_house_data" in st.session_state:
        with st.expander(t("last_details")):
            last_details = st.session_state["last_house_data"].iloc[0].to_dict()
            st.dataframe(
                pd.DataFrame([localize_property_details(last_details)]),
                width="stretch",
                hide_index=True
            )


# ==================== PAGES ====================
def render_languages_card():
    current = current_language()
    chips = "".join(
        f'<span class="lang-chip{" on" if name == current else ""}">{esc(name)}</span>' for name in TEXT.keys()
    )
    render_html(
        f'<div class="panel"><div class="section-title">{t("languages_title")}</div>'
        f'<div class="section-sub" style="margin-bottom:0">{t("languages_sub")}</div>'
        f'<div class="lang-row">{chips}</div></div>'
    )


def render_home_page(india_locations, area_type_options, india_model):
    render_hero()
    render_prediction_workspace(india_locations, area_type_options, india_model)
    render_divider()

    cards = (
        ("📄", "card_details_title", "card_details_text", "predict", "home_open_details"),
        ("🕘", "card_history_title", "card_history_text", "history", "home_open_history"),
        ("⬇️", "card_download_title", "card_download_text", "history", "home_open_download"),
    )
    for col, (icon, title_key, text_key, target, button_key) in zip(st.columns(3, gap="medium"), cards):
        with col:
            render_html(
                f'<div class="qcard"><div class="qcard-ico">{icon}</div><div>'
                f'<div class="qcard-title">{t(title_key)}</div>'
                f'<div class="qcard-text">{t(text_key)}</div></div></div>'
            )
            st.button(
                f'{t("open_btn")} →', key=button_key, width="stretch",
                on_click=go_to_page, args=(target,),
            )
    st.markdown('<div style="height:14px"></div>', unsafe_allow_html=True)
    render_languages_card()


def render_predict_page(india_locations, area_type_options, india_model):
    render_page_header(t("nav_predict"), t("india_sub"))
    render_prediction_workspace(india_locations, area_type_options, india_model)


def render_history_page():
    country_name_shown = t("country_India")
    render_page_header(
        t("history_title", country=country_name_shown),
        t("history_sub", country=country_name_shown),
    )

    history_df, load_error = load_prediction_history()
    if load_error:
        st.error(t("err_history_load", error=load_error))
    elif history_df.empty:
        render_html(
            f'<div class="empty-state"><div class="empty-ico">🏡</div>'
            f'<div class="empty-title">{t("history_empty_title")}</div>'
            f'<div class="empty-text">{t("history_empty", country=esc(country_name_shown))}</div></div>'
        )
        st.markdown('<div style="height:10px"></div>', unsafe_allow_html=True)
        _, button_col, _ = st.columns([1, 1, 1])
        with button_col:
            st.button(
                t("go_predict"), key="history_go_predict", width="stretch",
                on_click=go_to_page, args=("predict",),
            )
    else:
        render_html(f'<span class="count-pill">{esc(t("history_count", country=country_name_shown, count=len(history_df)))}</span>')
        readable_history = format_history_for_display(history_df).iloc[::-1]
        st.dataframe(
            readable_history,
            width="stretch",
            hide_index=True,
        )
        download_col, clear_col, _ = st.columns([1, 1, 2])
        with download_col:
            if EXCEL_AVAILABLE:
                st.download_button(
                    t("download_excel", country=country_name_shown),
                    data=history_to_excel_bytes(
                        readable_history, t("sheet_name", country=country_name_shown)
                    ),
                    file_name="india_prediction_history.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                    width="stretch",
                    key="download_history_India",
                )
            else:
                st.download_button(
                    t("download_csv", country=country_name_shown),
                    data=readable_history.to_csv(index=False).encode("utf-8-sig"),
                    file_name="india_prediction_history.csv",
                    mime="text/csv",
                    width="stretch",
                    key="download_history_India",
                )
                st.caption(t("excel_hint"))
        with clear_col:
            if st.button(
                t("clear_btn", country=country_name_shown),
                width="stretch",
                key="clear_history_India",
            ):
                clear_error = clear_prediction_history()
                if clear_error:
                    st.error(t("err_history_clear", error=clear_error))
                else:
                    st.success(t("history_cleared", country=country_name_shown))
                    st.rerun()


def render_profile_page():
    user_email = st.session_state["user"]["email"] or ""
    render_page_header(t("profile_title"), t("profile_sub"))
    account_col, about_col = st.columns(2, gap="large")
    with account_col:
        render_html(f"""
        <div class="panel">
          <div class="kv"><span class="k">{t('profile_email')}</span><span class="v">{esc(user_email)}</span></div>
          <div class="kv"><span class="k">{t('profile_language')}</span><span class="v">{esc(current_language())}</span></div>
          <div class="kv"><span class="k">{t('stat_dataset')}</span><span class="v">{t('dataset_india')}</span></div>
          <div class="kv"><span class="k">{t('stat_model')}</span><span class="v">{html_lib.escape(st.session_state.get('model_display_name', t('model_extra_trees')))}</span></div>
        </div>
        """)
    with about_col:
        render_html(
            f'<div class="panel"><div class="section-title">{t("profile_about_title")}</div>'
            f'<div class="section-sub" style="margin-bottom:0">{t("profile_about_text")}</div></div>'
        )
    st.markdown('<div style="height:14px"></div>', unsafe_allow_html=True)
    render_languages_card()


def render_footer():
    render_html(
        f'<div class="footer"><span>{t("footer_left")}</span>'
        f'<span>{t("footer_right")}</span></div>'
    )


# ==================== APP START ====================
inject_css()

# -------------------- LOGIN GATE --------------------
if "user" not in st.session_state:
    render_hero()
    show_login_screen()
    render_footer()
    st.stop()

# -------------------- LOAD MODEL AND DATA --------------------
try:
    india_model = load_model(INDIA_MODEL_PATH)
    st.session_state["model_display_name"] = model_display_name(india_model)
    area_type_options = load_area_types(INDIA_DATA_PATH)
    india_locations = load_india_locations(INDIA_DATA_PATH)
except Exception as exc:
    render_sidebar()
    st.error(t("err_model_load", name=Path(getattr(exc, "filename", "") or "model / dataset").name))
    with st.expander(t("tech_details")):
        st.exception(exc)
    st.stop()

# Clear any prediction state left by an earlier country-selection session.
if (
    st.session_state.get("prediction_country") not in (None, "India")
    or st.session_state.get("_last_country") not in (None, "India")
):
    clear_prediction_state()

# -------------------- NAVIGATION + PAGES --------------------
if st.session_state.get("page") not in PAGES:
    st.session_state["page"] = "home"
render_sidebar()

active_page = st.session_state["page"]
if active_page == "home":
    render_home_page(india_locations, area_type_options, india_model)
elif active_page == "predict":
    render_predict_page(india_locations, area_type_options, india_model)
elif active_page == "history":
    render_history_page()
else:
    render_profile_page()

render_footer()