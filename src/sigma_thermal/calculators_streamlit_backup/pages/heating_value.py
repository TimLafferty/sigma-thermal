"""
Heating Value Calculator

Calculate higher and lower heating values for gaseous fuels.
"""

import streamlit as st
import sys
from pathlib import Path

# Add parent directory to path for local imports
parent_dir = Path(__file__).parent.parent
if str(parent_dir) not in sys.path:
    sys.path.insert(0, str(parent_dir))

from utils.ui_components import (
    show_metric_card, show_comparison_result, show_input_validation,
    export_results_json, show_info_box, show_equation, format_number
)
from data.presets import FUEL_PRESETS, get_fuel_composition

# Import sigma_thermal modules
from sigma_thermal.combustion import (
    GasComposition,
    hhv_mass_gas,
    lhv_mass_gas,
    hhv_volume_gas,
    lhv_volume_gas
)


def show_heating_value_calculator():
    """Display heating value calculator page."""

    st.title("Heating Value Calculator")

    st.markdown("""
    Calculate the higher heating value (HHV) and lower heating value (LHV)
    for gaseous fuels on both mass and volume basis.
    """)

    # Show equations
    with st.expander("Theory & Equations", expanded=False):
        st.markdown("""
        ### Heating Values

        **Higher Heating Value (HHV):** Also called gross heating value or gross calorific value.
        Includes the latent heat of condensation of water vapor formed during combustion.

        **Lower Heating Value (LHV):** Also called net heating value or net calorific value.
        Does not include latent heat of water condensation (water leaves as vapor).
        """)

        show_equation(r"HHV - LHV = m_{H_2O} \times h_{fg}", "Difference")

        st.markdown("""
        Where:
        - $m_{H_2O}$ = mass of water formed (lb/lb fuel)
        - $h_{fg}$ = latent heat of vaporization (≈1050 BTU/lb at standard conditions)

        ### Typical Values

        | Fuel | HHV (BTU/lb) | LHV (BTU/lb) |
        |------|--------------|--------------|
        | Methane (CH4) | 23,875 | 21,495 |
        | Natural Gas | 22,000-23,000 | 20,000-21,000 |
        | Propane (C3H8) | 21,669 | 19,937 |
        | Hydrogen (H2) | 61,100 | 51,590 |
        """)

    # Preset selection
    st.subheader("Fuel Selection")

    col1, col2 = st.columns([2, 1])

    with col1:
        preset = st.selectbox(
            "Load Example Fuel:",
            list(FUEL_PRESETS.keys()),
            help="Select a preset fuel composition or choose 'Custom'"
        )

    with col2:
        if preset != "Custom":
            st.info(FUEL_PRESETS[preset]['description'])

    # Get preset composition
    if preset == "Custom":
        composition = {}
    else:
        composition = get_fuel_composition(preset)

    # Fuel composition inputs
    st.subheader("Fuel Composition (Mass %)")

    col1, col2, col3 = st.columns(3)

    with col1:
        st.markdown("**Paraffins**")
        ch4 = st.number_input(
            "Methane (CH4)",
            min_value=0.0,
            max_value=100.0,
            value=composition.get('methane_mass', 0.0),
            step=0.1,
            format="%.2f",
            help="CH4 content in mass %"
        )

        c2h6 = st.number_input(
            "Ethane (C2H6)",
            min_value=0.0,
            max_value=100.0,
            value=composition.get('ethane_mass', 0.0),
            step=0.1,
            format="%.2f"
        )

        c3h8 = st.number_input(
            "Propane (C3H8)",
            min_value=0.0,
            max_value=100.0,
            value=composition.get('propane_mass', 0.0),
            step=0.1,
            format="%.2f"
        )

        c4h10 = st.number_input(
            "Butane (C4H10)",
            min_value=0.0,
            max_value=100.0,
            value=composition.get('butane_mass', 0.0),
            step=0.1,
            format="%.2f"
        )

    with col2:
        st.markdown("**Other Combustibles**")
        h2 = st.number_input(
            "Hydrogen (H2)",
            min_value=0.0,
            max_value=100.0,
            value=composition.get('hydrogen_mass', 0.0),
            step=0.1,
            format="%.2f"
        )

        co = st.number_input(
            "Carbon Monoxide (CO)",
            min_value=0.0,
            max_value=100.0,
            value=composition.get('carbon_monoxide_mass', 0.0),
            step=0.1,
            format="%.2f"
        )

        h2s = st.number_input(
            "Hydrogen Sulfide (H2S)",
            min_value=0.0,
            max_value=100.0,
            value=composition.get('hydrogen_sulfide_mass', 0.0),
            step=0.1,
            format="%.2f",
            help="H2S content (if present)"
        )

    with col3:
        st.markdown("**Inerts**")
        co2 = st.number_input(
            "Carbon Dioxide (CO2)",
            min_value=0.0,
            max_value=100.0,
            value=composition.get('carbon_dioxide_mass', 0.0),
            step=0.1,
            format="%.2f",
            help="CO2 is inert (does not combust)"
        )

        n2 = st.number_input(
            "Nitrogen (N2)",
            min_value=0.0,
            max_value=100.0,
            value=composition.get('nitrogen_mass', 0.0),
            step=0.1,
            format="%.2f",
            help="N2 is inert (does not combust)"
        )

    # Validate total
    total = ch4 + c2h6 + c3h8 + c4h10 + h2 + co + h2s + co2 + n2

    st.markdown("---")

    # Check total is 100%
    valid = show_input_validation(total, 100.0, 0.01, "Total Composition")

    if not valid:
        st.stop()

    # Warn if high inerts
    inerts = co2 + n2
    if inerts > 20:
        st.warning(f"High inert content ({inerts:.1f}%). This will reduce heating value.")

    # Calculate button
    st.markdown("---")

    if st.button("Calculate Heating Values", type="primary", use_container_width=True):

        # Create fuel composition object
        try:
            fuel = GasComposition(
                methane_mass=ch4,
                ethane_mass=c2h6,
                propane_mass=c3h8,
                butane_mass=c4h10,
                hydrogen_mass=h2,
                carbon_monoxide_mass=co,
                hydrogen_sulfide_mass=h2s,
                carbon_dioxide_mass=co2,
                nitrogen_mass=n2
            )

            # Calculate heating values
            hhv_mass = hhv_mass_gas(fuel)
            lhv_mass = lhv_mass_gas(fuel)
            hhv_vol = hhv_volume_gas(fuel)
            lhv_vol = lhv_volume_gas(fuel)

            diff_mass = hhv_mass - lhv_mass
            diff_vol = hhv_vol - lhv_vol

            # Get decimal places from settings
            decimal_places = st.session_state.settings.get('decimal_places', 2)

            # Display results
            st.success("Calculation complete")

            st.subheader("Results")

            # Mass basis results
            st.markdown("### Mass Basis")

            col1, col2, col3 = st.columns(3)

            with col1:
                st.metric(
                    "Higher Heating Value",
                    f"{format_number(hhv_mass, 0, True)} BTU/lb",
                    help="HHV on mass basis"
                )

            with col2:
                st.metric(
                    "Lower Heating Value",
                    f"{format_number(lhv_mass, 0, True)} BTU/lb",
                    help="LHV on mass basis"
                )

            with col3:
                st.metric(
                    "Difference",
                    f"{format_number(diff_mass, 0, True)} BTU/lb",
                    help="HHV - LHV (latent heat of water)"
                )

            # Volume basis results
            st.markdown("### Volume Basis")

            col1, col2, col3 = st.columns(3)

            with col1:
                st.metric(
                    "HHV (Volume)",
                    f"{format_number(hhv_vol, 0, True)} BTU/scf",
                    help="HHV on volume basis (at standard conditions)"
                )

            with col2:
                st.metric(
                    "LHV (Volume)",
                    f"{format_number(lhv_vol, 0, True)} BTU/scf",
                    help="LHV on volume basis"
                )

            with col3:
                st.metric(
                    "Difference",
                    f"{format_number(diff_vol, 0, True)} BTU/scf",
                    help="Latent heat of water (volume basis)"
                )

            # Excel comparison (if known values)
            if preset == "Pure Methane":
                st.markdown("---")
                st.subheader("Excel VBA Comparison")

                st.markdown("**Mass Basis:**")
                show_comparison_result(
                    hhv_mass, 23875.0,
                    label="HHV",
                    unit="BTU/lb",
                    tolerance=0.01,
                    decimal_places=decimal_places
                )

            # Export options
            st.markdown("---")
            st.subheader("Export Results")

            results = {
                "Fuel Preset": preset,
                "Composition": {
                    "CH4 (%)": ch4,
                    "C2H6 (%)": c2h6,
                    "C3H8 (%)": c3h8,
                    "C4H10 (%)": c4h10,
                    "H2 (%)": h2,
                    "CO (%)": co,
                    "H2S (%)": h2s,
                    "CO2 (%)": co2,
                    "N2 (%)": n2
                },
                "Results (Mass Basis)": {
                    "HHV (BTU/lb)": round(hhv_mass, decimal_places),
                    "LHV (BTU/lb)": round(lhv_mass, decimal_places),
                    "Difference (BTU/lb)": round(diff_mass, decimal_places)
                },
                "Results (Volume Basis)": {
                    "HHV (BTU/scf)": round(hhv_vol, decimal_places),
                    "LHV (BTU/scf)": round(lhv_vol, decimal_places),
                    "Difference (BTU/scf)": round(diff_vol, decimal_places)
                }
            }

            col1, col2 = st.columns(2)

            with col1:
                export_results_json(results, "heating_value_results.json")

        except Exception as e:
            st.error(f"❌ Calculation Error: {str(e)}")
            st.exception(e)


if __name__ == "__main__":
    show_heating_value_calculator()
