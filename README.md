---
title: STAR Interview Simulator
emoji: "⭐"
colorFrom: blue
colorTo: yellow
sdk: streamlit
sdk_version: "1.32.0"
app_file: app.py
pinned: false
---

# STAR Interview Simulator

This Space runs the Streamlit app from `ai_star_interview_prep.py` via `app.py`.

## Required Secret

Set this in your Space **Settings -> Repository secrets**:

- `HF_TOKEN` = your Hugging Face token

The app uses this token to call the `microsoft/Phi-3-mini-4k-instruct` inference endpoint.
