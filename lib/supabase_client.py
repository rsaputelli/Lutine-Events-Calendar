import streamlit as st
from supabase import create_client, Client

@st.cache_resource
def get_public_supabase() -> Client:
    supa = st.secrets["supabase"]
    return create_client(supa["url"], supa["anon_key"])

@st.cache_resource
def get_admin_supabase() -> Client:
    supa = st.secrets["supabase"]
    return create_client(supa["url"], supa["key"])

def get_supabase():
    return get_admin_supabase()
