import streamlit as st
import io
import pandas as pd
import re
from datetime import date
from PyPDF2 import PdfReader
from docx import Document

# SharePoint SDK bits
from office365.sharepoint.client_context import ClientContext
from office365.sharepoint.files.file import File

# ---------- CONFIG ----------
SITE_URL = "https://eleven090.sharepoint.com/sites/Recruiting"
LIBRARY = "Shared Documents"
FOLDER = "Active Resumes"  # folder inside the doc library

# ---------- LOCAL COOKIE CONNECTOR ----------
# Reuse your signed-in browser session (MFA already done) when running locally.
# Requires: pip install browser-cookie3
import browser_cookie3

def _get_fedauth_rtfa():
    """Look up FedAuth/rtFa cookies from Chrome or Edge for *.sharepoint.com."""
    def pull(cj):
        fedauth = rtfa = None
        for c in cj:
            if c.domain.endswith("sharepoint.com"):
                n = c.name.lower()
                if n == "fedauth": fedauth = c.value
                elif n == "rtfa":  rtfa = c.value
        return fedauth, rtfa

    # Try Chrome
    try:
        cj = browser_cookie3.chrome(domain_name=".sharepoint.com")
        f, r = pull(cj)
        if f and r:
            return f, r
    except Exception:
        pass

    # Try Edge
    try:
        cj = browser_cookie3.edge(domain_name=".sharepoint.com")
        f, r = pull(cj)
        if f and r:
            return f, r
    except Exception:
        pass

    return None, None

def connect_with_browser_cookies():
    """
    Use your current browser session (MFA already done) to authenticate Office365 calls.
    Run this ONLY on your local machine (same OS user, no
