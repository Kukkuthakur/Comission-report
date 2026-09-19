# Commission Report Builder

Streamlit web app that turns a flat Excel file of patient tests into a
formatted commission report — index sheet + per-referrer statements.

## Formula
Rate = CutRate − Discount − Ambulance  (ambulance only when 100), floored at 0.

## Required input columns
PatientName, BillDate, ReferBy, TestName, PatientRate, DiscPercent, CutRate, Ambulance 
