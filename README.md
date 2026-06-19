# 📞 JD Agentic AI Call Evaluation System

An end-to-end pipeline that evaluates JustDial AI Bot call recordings — transcription, speaker identification, field verification, and call quality scoring — fully automated from MP3 to Excel.

**🔗 Live Demo:** [pridwimnjha-jd-call-evaluator.hf.space](https://pridwimnjha-jd-call-evaluator.hf.space)

---

## 🎯 Problem

JustDial's AI Bot makes thousands of calls daily to verify vendor business details (company name, address, mobile, working hours, contact person, business category). There was no automated way to:

- Check whether the bot followed the verification script
- Measure whether the vendor cooperated
- Extract structured data from the conversation
- Scale quality review across large call volumes

## ✅ Solution

Upload call recordings → get a fully structured 6-sheet Excel report in minutes. No manual listening required.

---

## 🏗️ Architecture

```
MP3 Upload
    ↓
AssemblyAI (transcription + speaker diarization + language detection)
    ↓
3-Layer Speaker Identification (Bot vs Vendor)
    ↓
Hinglish Transliteration (Devanagari → readable Roman script)
    ↓
Programmatic Metrics (latency, on-script ratio, repeat count)
    ↓
Groq LLaMA 3.3 70B (structured JSON evaluation)
    ↓
Schema Validation + Hallucination Guard
    ↓
Manual Review Flagging
    ↓
6-Sheet Excel Report
```

---

## ✨ Features

### 1. Transcription
Converts MP3 call recordings into readable Hinglish (Hindi written in Roman script) transcripts with speaker labels.

### 2. Speaker Identification (3-Layer)
- **Layer 1** — Keyword scoring (Bot-specific phrases like "verification", "confirm")
- **Layer 2** — First-utterance length heuristic (handles vendor picking up first)
- **Layer 3** — Vendor-phrase safety net (flips mapping if needed)

Handles all edge cases including vendor speaking first, single-speaker calls, and crosstalk.

### 3. Hinglish Transliteration
Custom-built character mapping (not a generic library) converts Devanagari to natural, readable Hinglish:
```
हेलो वेरिफिकेशन कॉल है  →  hello verification call hai
```
Includes nukta-character handling (ज़, फ़) and common English-loanword fixes (`veriphikeshn → verification`).

### 4. Field Verification
Extracts and verifies 6 business fields:

| Field | Status Options |
|---|---|
| Company Name | Confirmed / Partial / Not Confirmed |
| Address | Confirmed / Partial / Not Confirmed |
| Mobile Number | Confirmed / Partial / Not Confirmed |
| Working Hours | Confirmed / Partial / Not Confirmed |
| Contact Person | Confirmed / Partial / Not Confirmed |
| Business Category | Confirmed / Partial / Not Confirmed |

Every extraction includes a confidence score and a verbatim transcript quote for traceability.

### 5. Bot Quality Scoring
Weighted 4-dimension rubric:

| Dimension | Weight |
|---|---|
| Script adherence | 30% |
| Question relevance | 30% |
| No repetition | 20% |
| Professionalism | 20% |

Explicitly judges bot performance **independent of vendor cooperation** — a bot that did everything right but was refused by an uncooperative vendor still scores Good.

### 6. User Quality Scoring
Evaluates vendor cooperation, clarity, frustration, and engagement.

### 7. Data Collection Score
A separate end-result metric (Excellent / Good / Partial / Failed) based purely on how many fields were collected — independent of Bot Quality. This separates *"did the bot perform well"* from *"was data actually collected."*

### 8. Sentiment Analysis
Vendor sentiment: Positive / Neutral / Negative.

### 9. Programmatic Metrics (No AI guessing)
Computed directly from AssemblyAI timestamps and text:
- **First response latency** — seconds before vendor first responds
- **On-script ratio** — fraction of bot turns matching ~50 verification keywords
- **Repeat question count** — `SequenceMatcher` similarity detection
- **Token cost estimate** — USD cost per call

### 10. Hallucination Guard
Every "Confirmed" field is checked against the actual transcript using fuzzy 70% word-overlap matching. If the AI's quote doesn't exist in the transcript, confidence is automatically reduced and the call is flagged for manual review.

### 11. Manual Review Detection
Auto-flags calls needing human review:
- Call duration < 30 seconds
- More than 2 speakers detected
- Off-script ratio below threshold
- Repeated questions (≥2)
- Low-confidence field extractions

### 12. Excel Report (6 Sheets)

| Sheet | Contents |
|---|---|
| **Summary** | One row per call — all scores + full transcript |
| **Fields Detail** | Per-field status, value, confidence, transcript quote |
| **Metrics** | Latency, on-script ratio, repeat count, cost |
| **Transcripts** | Full conversation text |
| **Raw JSON** | Complete LLM output for debugging |
| **Accuracy** | Ground-truth comparison template for manual QA |

---

## 🧠 Tech Stack

| Component | Technology |
|---|---|
| Transcription & Diarization | [AssemblyAI](https://www.assemblyai.com/) |
| LLM Evaluation | [Groq](https://groq.com/) — LLaMA 3.3 70B |
| Frontend | [Streamlit](https://streamlit.io/) |
| Excel Generation | `pandas` + `openpyxl` |
| Deployment | [Hugging Face Spaces](https://huggingface.co/spaces) |

---

## 💰 Cost

| Scale | Estimated Cost |
|---|---|
| 5 calls (pilot) | ~$0.0006 |
| 1,000 calls | ~$0.12 (~₹10) |
| 10,000 calls | ~$1.18 (~₹98) |
| 1,00,000 calls | ~$11.80 (~₹980) |

Both AssemblyAI and Groq free tiers are generous enough that this entire pipeline runs near-zero cost even at scale:
- AssemblyAI: 100 hours/month free
- Groq: 14,400 requests/day free, never expires

---

## 🚀 Setup

### Prerequisites
- Python 3.10+
- AssemblyAI API key ([get one free](https://www.assemblyai.com/))
- Groq API key ([get one free](https://console.groq.com/))

### Local Installation

```bash
git clone https://github.com/<your-username>/jd-call-evaluator.git
cd jd-call-evaluator
pip install -r requirements.txt
```

### Environment Variables

```bash
export ASSEMBLYAI_API_KEY="your_key_here"
export GROQ_API_KEY="your_key_here"
```

### Run Locally

```bash
streamlit run app.py
```

### Deploy on Hugging Face Spaces

1. Create a new Space → SDK: **Streamlit**
2. Upload `app.py` and `requirements.txt`
3. Add `ASSEMBLYAI_API_KEY` and `GROQ_API_KEY` under **Settings → Repository Secrets**
4. Space builds and deploys automatically

---

## 📊 Sample Output

```
File: G31Z9G8V3.mp3
Duration: 69s | Language: Hindi | Latency: 1.16s | Cost: $0.000115

Bot Quality:            Good
User Quality:           Poor
Sentiment:              Negative
Call Outcome:           Failed
Data Collection Score:  Failed
On-Script Ratio:        0.67

Transcript (Hinglish):
Bot:    jaldi se neha bol rhi hun ye verification kaun hai
        bs aapke business ki jankari confirm karni hai
Vendor: han shi hai lekin kal fon aaya main kiya tha
Vendor: verification ho gya jo sale se mer ko verification mail bhi aaya
Bot:    sirf 1 mint lagega bs kuch besik details confirm karni hai
```

---

## ⚠️ Known Limitations

- Off-script detection is keyword-based, not semantic — may need keyword list expansion as more call patterns appear
- Transliteration has minor edge cases for complex Devanagari conjuncts
- Recommended max batch size: ~50 files per run on Hugging Face free tier (sequential processing, no checkpointing yet)
- Ground-truth accuracy not yet measured at scale — Accuracy sheet provided for manual QA

## 🔮 Future Improvements

- Batch processing with checkpointing for 1000+ call datasets
- Semantic (embedding-based) off-script detection
- Real-time call monitoring dashboard
- Database integration for historical call tracking
- Multi-language support (Tamil, Telugu, Kannada)
- Fine-tuned evaluation model on JD-specific call data

---

## 👨‍💻 Author

**Pridwimn Sanjeeb Jha**
Business Analyst Intern, JustDial Ltd. — JD Mart
B.Tech CS + Data Science, NMIMS Mumbai

[LinkedIn](https://linkedin.com/in/pridwimnjha) · [Live Demo](https://pridwimnjha-jd-call-evaluator.hf.space)

---

## 📄 License

This project was built as part of an internship deliverable for JustDial Ltd. All rights reserved by JustDial Ltd. unless otherwise specified.
