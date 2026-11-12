<div align="center">

# 🏡 BLR HOME IQ

### AI-Powered House Price Estimation for Bengaluru

**A multilingual, interactive property valuation web app built with Python, Scikit-learn, Streamlit and Supabase.**

`Machine Learning` · `Real Estate` · `Streamlit` · `Supabase` · `Python`

</div>

---

## 📌 Project Overview

**BLR HOME IQ** estimates the sale price of residential properties in **Bengaluru, India** using a trained machine-learning model and six property characteristics. Users can create an account, enter property details, view a price estimate in Indian rupees, and manage their saved prediction history.

This is a **Bengaluru-specific educational project**, not an all-India pricing engine or an official property valuation service.

## ✨ Key Features

| Feature | Description |
|---|---|
| 🧠 ML-based valuation | Generates property price estimates using a saved Scikit-learn model |
| 📍 Bengaluru locality selection | Choose a locality from the project's dataset |
| 🏠 Property input form | Enter six characteristics of a home |
| 💰 Indian price formatting | Displays estimates in INR lakh / crore |
| 🔐 Authentication | Account signup and login powered by Supabase |
| 🕘 Prediction history | Save, view and clear estimates in the user's account |
| 📥 Export | Download prediction history as CSV or Excel |
| 🌐 Multilingual UI | English and multiple Indian-language interface options |
| 📱 Responsive dashboard | Streamlit interface with Home, Predict, History and Profile pages |

## 🧾 Model Inputs

The prediction uses these **six** features:

| Input | Meaning |
|---|---|
| `location` | Bengaluru locality |
| `area_type` | Area classification from the property listing |
| `bhk` | Number of bedrooms |
| `total_sqft` | Property area in square feet |
| `bath` | Number of bathrooms |
| `balcony` | Number of balconies |

**Output:** Estimated property price in **₹ lakh** (also formatted as crore when appropriate).

## 🛠️ Technology Stack

- **Frontend and app:** Streamlit, HTML/CSS embedded in Python
- **Programming and data processing:** Python, Pandas, NumPy
- **Machine learning:** Scikit-learn, Joblib
- **Authentication and cloud history:** Supabase
- **Export:** CSV and OpenPyXL for Excel files
- **Version control:** Git and GitHub

## 🧠 Machine Learning Approach

The project explores tree-based regression models, including **Random Forest** and **Extra Trees**, and evaluates their predictions against held-out data. The app loads its deployed estimator from `india_house_price_model.pkl`; the exact algorithm displayed by the app is detected from that saved estimator.

### Evaluation on previously unseen locations

The following experimental results come from a **location-disjoint holdout test** (2,310 test rows). These are experiment results and **do not automatically certify the model currently saved in the `.pkl` file**.

| Experimental model | MAE (₹ lakh) | R² |
|---|---:|---:|
| Random Forest + median bias correction | 32.01 | 0.6943 |
| Extra Trees + median bias correction | 32.09 | 0.6707 |
| Log-target Random Forest + correction | 31.88 | 0.6322 |

**Interpretation:** These tests are intentionally challenging because test locations are excluded from model training. Predictions can be less reliable for unfamiliar localities and unusually expensive properties. The log-target experiment has slightly lower MAE but noticeably lower R²; lower MAE alone does not make it the best model for every use case.

## 📁 Project Structure

```text
BLR-HOME-IQ/
├── app.py                              # Streamlit application
├── india_house_price_model.pkl        # Saved trained estimator
├── DATA/
│   ├── Bengaluru_House_Data_Final.csv  # Final project dataset
│   └── Bengaluru_House_Data_Cleaned.csv # Compatible cleaned dataset
├── requirements.txt                   # Python dependencies
├── README.md                          # Project documentation
└── .gitignore                         # Excludes local/private files
```

The app prefers `DATA/Bengaluru_House_Data_Final.csv` and falls back to the cleaned dataset if needed. The saved model and CSV should correspond to the same preprocessing and feature definitions.

## 🚀 Run Locally

### 1. Clone the repository

```bash
git clone https://github.com/YOUR_USERNAME/BLR-HOME-IQ.git
cd BLR-HOME-IQ
```

Replace `YOUR_USERNAME` with your GitHub username.

### 2. Create and activate a virtual environment

**Windows PowerShell:**

```powershell
python -m venv .venv
.\.venv\Scripts\Activate.ps1
```

### 3. Install dependencies

```bash
pip install -r requirements.txt
```

### 4. Configure Supabase

Create a **local, untracked** file at `.streamlit/secrets.toml`:

```toml
SUPABASE_URL = "YOUR_SUPABASE_PROJECT_URL"
SUPABASE_KEY = "YOUR_SUPABASE_PUBLIC_ANON_KEY"
```

Use only the **public/publishable (anon) key** in the app. **Never publish a Supabase service-role or secret key.** The Supabase project must already have the required authentication and `prediction_history` table configured, with Row Level Security policies that restrict each user to their own records. The README does not create that database table automatically.

### 5. Launch

```bash
streamlit run app.py
```

Open the local URL displayed by Streamlit in your browser.

## ☁️ Deployment

For Streamlit Community Cloud:

1. Push the safe project files to a GitHub repository.
2. Open [Streamlit Community Cloud](https://share.streamlit.io/) and create an app from that repository.
3. Set the entry point to `app.py`.
4. Configure `SUPABASE_URL` and the public/anon `SUPABASE_KEY` in the deployment's **Secrets** settings, not in GitHub.
5. Deploy and test signup, login, prediction and history.

> Deployment has **not** been claimed as complete in this repository documentation; the steps above describe how to deploy it.

## 🔒 Security and Privacy

- Do **not** commit `.env`, `.streamlit/secrets.toml`, passwords, private keys, or `.venv/`.
- Do **not** upload `DATA/india_prediction_history.csv` or other files containing users' personal history.
- Enable and verify **Supabase Row Level Security** for user-specific history access.
- Load `.pkl` / Joblib model files only from trusted sources.
- Property estimates are approximate and should not be used as the sole basis for financial decisions.

## ⚠️ Limitations

- Covers **Bengaluru localities**, not properties across India.
- Price estimates depend on the dataset and its coverage.
- Location-based holdout testing indicates meaningful error, especially for high-value properties.
- The model does not directly include all possible price drivers, such as live market conditions, property condition or legal status.
- A predicted value is **not** a certified appraisal or a guaranteed selling price.

## 🧭 Future Improvements

- Model monitoring and updated market data
- Better performance on premium properties and less-represented localities
- Additional validated property features where reliable data is available
- Improved model evaluation and documented deployment checks

## 📄 License

No open-source license has been selected yet. Add a `LICENSE` file if you decide to publish this project under an open-source license.

---

<div align="center">

**BLR HOME IQ** · Bengaluru Property Intelligence

*Built as a machine-learning portfolio project.*

</div>
