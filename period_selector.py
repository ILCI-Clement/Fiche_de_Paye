"""Named French month selectors with synchronized controls."""
import streamlit as st

MONTH_NAMES = ("Janvier", "Février", "Mars", "Avril", "Mai", "Juin",
               "Juillet", "Août", "Septembre", "Octobre", "Novembre", "Décembre")


def select_period(scope, location, month, year):
    period_key = f"period_{scope}"
    st.session_state.setdefault(period_key, (int(month), int(year)))
    month_key = f"{period_key}_{location}_month"
    year_key = f"{period_key}_{location}_year"
    selected_month, selected_year = st.session_state[period_key]
    st.session_state[month_key] = selected_month
    st.session_state[year_key] = selected_year

    def sync():
        st.session_state[period_key] = (st.session_state[month_key], st.session_state[year_key])
        if scope.startswith("fiches_"):
            for key in list(st.session_state):
                if key.startswith("generated_fiche_"):
                    del st.session_state[key]

    left, right = st.columns(2)
    with left:
        st.selectbox("Mois", range(1, 13), format_func=lambda value: MONTH_NAMES[value - 1],
                     key=month_key, on_change=sync)
    with right:
        st.number_input("Année", min_value=2000, max_value=2100, key=year_key, on_change=sync)
    return st.session_state[period_key]
