"""
Sigma Thermal Engineering Calculators

A Streamlit-based web application providing engineering calculation tools
for combustion, fluids, and thermal systems.

Author: GTS Energy Inc.
Date: October 2025
"""

import streamlit as st

# Configure page
st.set_page_config(
    page_title="Sigma Thermal",
    page_icon="⚡",
    layout="wide",
    initial_sidebar_state="expanded"
)

# Custom CSS - Professional, minimal styling
st.markdown("""
<style>
    /* Typography */
    @import url('https://fonts.googleapis.com/css2?family=Inter:wght@300;400;500;600;700&display=swap');

    html, body, [class*="css"] {
        font-family: 'Inter', -apple-system, BlinkMacSystemFont, 'Segoe UI', sans-serif;
    }

    /* Sidebar styling */
    [data-testid="stSidebar"] {
        background-color: #f8f9fa;
        border-right: 1px solid #e9ecef;
    }

    /* Main content area */
    .main .block-container {
        padding-top: 2rem;
        padding-bottom: 2rem;
        max-width: 1200px;
    }

    /* Headers */
    h1 {
        font-weight: 600;
        color: #212529;
        margin-bottom: 1rem;
        font-size: 2rem;
    }

    h2 {
        font-weight: 600;
        color: #495057;
        margin-top: 2rem;
        margin-bottom: 1rem;
        font-size: 1.5rem;
    }

    h3 {
        font-weight: 500;
        color: #6c757d;
        margin-top: 1.5rem;
        margin-bottom: 0.75rem;
        font-size: 1.125rem;
    }

    /* Remove default streamlit styling */
    #MainMenu {visibility: hidden;}
    footer {visibility: hidden;}

    /* Buttons */
    .stButton > button {
        background-color: #0d6efd;
        color: white;
        border: none;
        border-radius: 4px;
        padding: 0.5rem 1.5rem;
        font-weight: 500;
        transition: all 0.2s;
    }

    .stButton > button:hover {
        background-color: #0b5ed7;
        box-shadow: 0 2px 4px rgba(0,0,0,0.1);
    }

    /* Input fields */
    .stNumberInput > div > div > input,
    .stTextInput > div > div > input,
    .stSelectbox > div > div > select {
        border-radius: 4px;
        border: 1px solid #dee2e6;
    }

    /* Cards */
    .metric-container {
        background: white;
        border: 1px solid #e9ecef;
        border-radius: 8px;
        padding: 1.25rem;
        margin: 0.5rem 0;
    }

    /* Dividers */
    hr {
        margin: 2rem 0;
        border: none;
        border-top: 1px solid #e9ecef;
    }

    /* Tables */
    .dataframe {
        border: 1px solid #e9ecef !important;
        border-radius: 4px;
    }

    /* Expander */
    .streamlit-expanderHeader {
        font-weight: 500;
        color: #495057;
    }
</style>
""", unsafe_allow_html=True)


def main():
    """Main application entry point."""

    # Sidebar navigation
    st.sidebar.title("Sigma Thermal")
    st.sidebar.markdown("---")

    page = st.sidebar.radio(
        "Navigate to:",
        [
            "Home",
            "Heating Value Calculator",
            "Air Requirement Calculator",
            "Products of Combustion",
            "Flue Gas Enthalpy",
            "Combustion Efficiency",
            "Steam Properties",
            "Water Properties",
            "Flash Steam Calculator",
            "Excel Comparison Tool",
        ]
    )

    # Settings in sidebar
    st.sidebar.markdown("---")
    st.sidebar.subheader("Settings")

    # Unit system (future)
    unit_system = st.sidebar.selectbox(
        "Unit System",
        ["US Customary", "SI (Metric)"],
        disabled=True,
        help="SI units coming soon"
    )

    # Precision
    decimal_places = st.sidebar.slider(
        "Decimal Places",
        min_value=0,
        max_value=6,
        value=2,
        help="Number of decimal places in results"
    )

    # Store settings in session state
    if 'settings' not in st.session_state:
        st.session_state.settings = {}

    st.session_state.settings['unit_system'] = unit_system
    st.session_state.settings['decimal_places'] = decimal_places

    # Display selected page
    if page == "Home":
        show_home()
    elif page == "Heating Value Calculator":
        from pages.heating_value import show_heating_value_calculator
        show_heating_value_calculator()
    elif page == "Air Requirement Calculator":
        from pages.air_requirement import show_air_requirement_calculator
        show_air_requirement_calculator()
    elif page == "Products of Combustion":
        from pages.products_combustion import show_products_combustion_calculator
        show_products_combustion_calculator()
    elif page == "Flue Gas Enthalpy":
        from pages.flue_gas_enthalpy import show_flue_gas_enthalpy_calculator
        show_flue_gas_enthalpy_calculator()
    elif page == "Combustion Efficiency":
        from pages.combustion_efficiency import show_combustion_efficiency_calculator
        show_combustion_efficiency_calculator()
    elif page == "Steam Properties":
        from pages.steam_properties import show_steam_properties_calculator
        show_steam_properties_calculator()
    elif page == "Water Properties":
        from pages.water_properties import show_water_properties_calculator
        show_water_properties_calculator()
    elif page == "Flash Steam Calculator":
        from pages.flash_steam import show_flash_steam_calculator
        show_flash_steam_calculator()
    elif page == "Excel Comparison Tool":
        from pages.excel_comparison import show_excel_comparison_tool
        show_excel_comparison_tool()


def show_home():
    """Display home page."""

    # Header
    st.title("Sigma Thermal Engineering Calculators")

    st.markdown("""
    Welcome to the Sigma Thermal calculator suite! This application provides comprehensive
    engineering calculations for combustion, fluids, and thermal systems.
    """)

    # Quick stats
    col1, col2, col3, col4 = st.columns(4)

    with col1:
        st.metric("Functions", "43", help="Total implemented functions")

    with col2:
        st.metric("Calculators", "9", help="Available calculator tools")

    with col3:
        st.metric("Tests", "412", help="Automated test coverage")

    with col4:
        st.metric("Accuracy", "<1%", help="Deviation from reference data")

    # Combustion Calculators
    st.markdown("---")
    st.subheader("Combustion Calculators")

    col1, col2 = st.columns(2)

    with col1:
        st.markdown("**Heating Value Calculator**")
        st.markdown("""
        Calculate higher and lower heating values for gaseous and liquid fuels.
        - Mass and volume basis
        - Multiple fuel components
        - Comparison to Excel VBA
        """)

        st.markdown("**Products of Combustion**")
        st.markdown("""
        Calculate flue gas composition from fuel and excess air.
        - H2O, CO2, N2, O2, SO2
        - Mass and volume basis
        - Composition charts
        """)

        st.markdown("**Combustion Efficiency**")
        st.markdown("""
        Analyze combustion system efficiency and losses.
        - Stack loss calculation
        - Radiation and other losses
        - Energy flow diagrams
        """)

    with col2:
        st.markdown("**Air Requirement Calculator**")
        st.markdown("""
        Determine stoichiometric and actual air requirements.
        - Theoretical air needed
        - Excess air effects
        - Air-fuel ratios
        """)

        st.markdown("**Flue Gas Enthalpy**")
        st.markdown("""
        Calculate flue gas energy content and sensible heat.
        - Component enthalpies
        - Temperature effects
        - Total enthalpy
        """)


    # Fluids Calculators
    st.markdown("---")
    st.subheader("Fluids & Steam Calculators")

    col1, col2 = st.columns(2)

    with col1:
        st.markdown("**Steam Properties**")
        st.markdown("""
        Calculate steam and water thermodynamic properties.
        - Saturation properties
        - Enthalpy and quality
        - Phase determination
        - T-s and P-h diagrams
        """)

        st.markdown("**Flash Steam Calculator**")
        st.markdown("""
        Analyze flash steam generation from pressure reduction.
        - Flash steam flow
        - Energy recovery
        - Economic analysis
        """)

    with col2:
        st.markdown("**Water Properties**")
        st.markdown("""
        Determine water transport and thermal properties.
        - Density and viscosity
        - Specific heat and conductivity
        - Prandtl number
        - Reynolds number calculator
        """)

    # Validation Tools
    st.markdown("---")
    st.subheader("Validation & Comparison")

    st.markdown("**Excel VBA Comparison Tool**")
    st.markdown("""
    Validate Python calculations against Excel VBA macros.
    - Side-by-side comparison
    - Deviation analysis
    - Batch testing
    - Discrepancy reporting
    """)

    # Footer
    st.markdown("---")
    st.markdown("""
    <div style="text-align: center; color: #7f8c8d; padding: 2rem;">
        <p><b>Sigma Thermal Engineering Calculators</b> v1.0</p>
        <p>Powered by Python • 43 functions • 412 tests • <1% accuracy</p>
        <p>© 2025 GTS Energy Inc.</p>
    </div>
    """, unsafe_allow_html=True)


if __name__ == "__main__":
    main()
