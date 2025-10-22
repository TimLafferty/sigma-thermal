"""
Steam Properties Calculator

Calculate steam and water thermodynamic properties including
saturation, enthalpy, and quality.
"""

import streamlit as st
import sys
from pathlib import Path
import plotly.graph_objects as go

# Add parent directory to path for local imports
parent_dir = Path(__file__).parent.parent
if str(parent_dir) not in sys.path:
    sys.path.insert(0, str(parent_dir))

from utils.ui_components import (
    show_metric_card, show_comparison_result, show_info_box,
    export_results_json, format_number, show_equation
)
from data.presets import STEAM_PRESSURE_PRESETS

# Import sigma_thermal modules
from sigma_thermal.fluids import (
    saturation_pressure,
    saturation_temperature,
    steam_enthalpy,
    steam_quality
)


def show_steam_properties_calculator():
    """Display steam properties calculator page."""

    st.title("Steam Properties Calculator")

    st.markdown("""
    Calculate thermodynamic properties of water and steam including saturation
    properties, enthalpy, and quality for any pressure and temperature.
    """)

    # Theory
    with st.expander("Theory & Equations", expanded=False):
        st.markdown("""
        ### Saturation Properties

        At any given pressure, water/steam exists at a unique saturation temperature.
        """)

        show_equation(r"P_{sat} = f(T)", "Saturation Pressure")
        show_equation(r"T_{sat} = f(P)", "Saturation Temperature")

        st.markdown("""
        ### Steam Enthalpy

        Enthalpy depends on temperature, pressure, and phase:
        """)

        show_equation(r"h = h_f + x \cdot h_{fg}", "Two-Phase Mixture")

        st.markdown("""
        Where:
        - $h_f$ = saturated liquid enthalpy (BTU/lb)
        - $h_{fg}$ = enthalpy of vaporization (BTU/lb)
        - $x$ = quality (0 = liquid, 1 = vapor)

        ### Steam Quality

        Quality is the mass fraction of vapor in a two-phase mixture:
        """)

        show_equation(r"x = \\frac{h - h_f}{h_{fg}}", "Quality from Enthalpy")

        st.markdown("""
        - $x < 0$: Subcooled liquid
        - $0 \\leq x \\leq 1$: Two-phase mixture
        - $x > 1$: Superheated vapor
        """)

    # Calculation mode selection
    st.subheader("Calculation Mode")

    calc_mode = st.radio(
        "Select input parameters:",
        [
            "Temperature & Pressure (Known)",
            "Enthalpy & Pressure (Known)",
            "Saturation Properties Only"
        ],
        horizontal=True
    )

    st.markdown("---")

    # Mode 1: T, P known
    if calc_mode == "Temperature & Pressure (Known)":

        st.subheader("📝 Input Parameters")

        col1, col2, col3 = st.columns(3)

        with col1:
            temperature = st.number_input(
                "Temperature (°F)",
                min_value=32.0,
                max_value=700.0,
                value=212.0,
                step=1.0,
                help="Water/steam temperature"
            )

        with col2:
            # Pressure preset selector
            pressure_preset = st.selectbox(
                "Pressure Preset",
                ["Custom"] + list(STEAM_PRESSURE_PRESETS.keys()),
                help="Select common pressure or enter custom"
            )

            if pressure_preset == "Custom":
                pressure = st.number_input(
                    "Pressure (psia)",
                    min_value=0.1,
                    max_value=3000.0,
                    value=14.7,
                    step=0.1,
                    help="System pressure"
                )
            else:
                pressure = STEAM_PRESSURE_PRESETS[pressure_preset]
                st.info(f"Pressure: {pressure} psia")

        with col3:
            quality = st.slider(
                "Quality (for two-phase)",
                min_value=0.0,
                max_value=1.0,
                value=1.0,
                step=0.01,
                help="0 = saturated liquid, 1 = saturated vapor"
            )

        # Calculate button
        if st.button("Calculate Steam Properties", type="primary", use_container_width=True):

            try:
                # Get saturation properties
                t_sat = saturation_temperature(pressure)
                p_sat = saturation_pressure(temperature)

                # Determine phase
                if temperature < t_sat - 0.5:
                    phase = "Subcooled Liquid"
                    phase_color = "#3498db"
                elif temperature > t_sat + 0.5:
                    phase = "Superheated Vapor"
                    phase_color = "#e74c3c"
                else:
                    phase = "Saturated (Two-Phase)"
                    phase_color = "#f39c12"

                # Calculate enthalpy
                enthalpy = steam_enthalpy(temperature, pressure, quality)

                # Calculate quality from enthalpy
                calc_quality = steam_quality(enthalpy, pressure)

                # Get saturation enthalpies
                hf = steam_enthalpy(t_sat, pressure, 0.0)
                hg = steam_enthalpy(t_sat, pressure, 1.0)
                hfg = hg - hf

                # Display results
                st.success("Calculation complete")

                # Phase indicator
                st.markdown(f"""
                <div style="background-color: {phase_color}; color: white;
                            padding: 1rem; border-radius: 4px; text-align: center;
                            font-size: 1.25rem; font-weight: 500; margin: 1rem 0;">
                    {phase}
                </div>
                """, unsafe_allow_html=True)

                st.subheader("Results")

                # Main properties
                col1, col2, col3 = st.columns(3)

                with col1:
                    st.metric(
                        "Enthalpy",
                        f"{format_number(enthalpy, 1)} BTU/lb",
                        help="Specific enthalpy at given conditions"
                    )

                with col2:
                    if -0.01 < calc_quality < 1.01:
                        st.metric(
                            "Quality",
                            f"{format_number(calc_quality * 100, 1)} %",
                            help="Vapor mass fraction (0-100%)"
                        )
                    else:
                        st.metric(
                            "Quality",
                            "N/A",
                            help="Single phase (subcooled or superheated)"
                        )

                with col3:
                    st.metric(
                        "Saturation Temp",
                        f"{format_number(t_sat, 1)} °F",
                        help=f"Tsat at {pressure} psia"
                    )

                # Saturation properties table
                st.markdown("### Saturation Properties")

                sat_props = {
                    "Saturation Temperature": f"{format_number(t_sat, 2)} °F",
                    "Saturation Pressure (at T)": f"{format_number(p_sat, 2)} psia",
                    "Liquid Enthalpy (hf)": f"{format_number(hf, 1)} BTU/lb",
                    "Vapor Enthalpy (hg)": f"{format_number(hg, 1)} BTU/lb",
                    "Enthalpy of Vaporization (hfg)": f"{format_number(hfg, 1)} BTU/lb"
                }

                for prop, value in sat_props.items():
                    col1, col2 = st.columns([2, 1])
                    col1.markdown(f"**{prop}:**")
                    col2.markdown(value)

                # T-s diagram (simplified representation)
                st.markdown("### 📈 T-s Diagram")

                show_info_box(
                    "Current state marked on simplified temperature-entropy diagram",
                    "info"
                )

                # Create simplified T-s diagram
                fig = go.Figure()

                # Saturation line (simplified)
                temp_sat_line = [32, 100, 212, 300, 400, 500, 600, 700]
                s_f = [0.0, 0.13, 0.31, 0.43, 0.55, 0.65, 0.74, 0.82]  # Approximate
                s_g = [2.19, 1.98, 1.76, 1.63, 1.52, 1.44, 1.37, 1.30]  # Approximate

                fig.add_trace(go.Scatter(
                    x=s_f,
                    y=temp_sat_line,
                    mode='lines',
                    name='Saturated Liquid Line',
                    line=dict(color='blue', width=2)
                ))

                fig.add_trace(go.Scatter(
                    x=s_g,
                    y=temp_sat_line,
                    mode='lines',
                    name='Saturated Vapor Line',
                    line=dict(color='red', width=2)
                ))

                # Mark current state (approximate entropy)
                if calc_quality >= 0 and calc_quality <= 1:
                    # Two-phase: interpolate between sf and sg
                    import numpy as np
                    s_point = np.interp(t_sat, temp_sat_line, s_f) + \
                              calc_quality * (np.interp(t_sat, temp_sat_line, s_g) -
                                            np.interp(t_sat, temp_sat_line, s_f))
                    t_point = t_sat
                else:
                    # Single phase: approximate
                    s_point = 1.0  # Placeholder
                    t_point = temperature

                fig.add_trace(go.Scatter(
                    x=[s_point],
                    y=[t_point],
                    mode='markers',
                    name='Current State',
                    marker=dict(color=phase_color, size=15, symbol='star')
                ))

                fig.update_layout(
                    title="Temperature-Entropy Diagram (Simplified)",
                    xaxis_title="Entropy (BTU/lb·°R) - Approximate",
                    yaxis_title="Temperature (°F)",
                    height=500,
                    showlegend=True,
                    hovermode='closest'
                )

                st.plotly_chart(fig, use_container_width=True)

                # Validation (if at 14.7 psia)
                if abs(pressure - 14.7) < 0.1 and abs(temperature - 212.0) < 0.1:
                    st.markdown("---")
                    st.subheader("ASME Steam Table Comparison")

                    if quality == 0.0:
                        show_comparison_result(
                            hf, 180.1,
                            label="hf",
                            unit="BTU/lb",
                            tolerance=0.01
                        )
                    elif quality == 1.0:
                        show_comparison_result(
                            hg, 1150.4,
                            label="hg",
                            unit="BTU/lb",
                            tolerance=0.01
                        )

                # Export results
                st.markdown("---")
                st.subheader("Export Results")

                results = {
                    "Inputs": {
                        "Temperature (°F)": temperature,
                        "Pressure (psia)": pressure,
                        "Quality": quality
                    },
                    "Phase": phase,
                    "Results": {
                        "Enthalpy (BTU/lb)": round(enthalpy, 2),
                        "Calculated Quality": round(calc_quality, 4) if -0.01 < calc_quality < 1.01 else None,
                        "Saturation Temperature (°F)": round(t_sat, 2),
                        "Saturation Pressure at T (psia)": round(p_sat, 2),
                        "hf (BTU/lb)": round(hf, 2),
                        "hg (BTU/lb)": round(hg, 2),
                        "hfg (BTU/lb)": round(hfg, 2)
                    }
                }

                export_results_json(results, "steam_properties_results.json")

            except Exception as e:
                st.error(f"Calculation Error: {str(e)}")
                st.exception(e)

    # Mode 2: h, P known
    elif calc_mode == "Enthalpy & Pressure (Known)":

        st.subheader("📝 Input Parameters")

        col1, col2 = st.columns(2)

        with col1:
            enthalpy_input = st.number_input(
                "Enthalpy (BTU/lb)",
                min_value=0.0,
                max_value=1500.0,
                value=665.0,
                step=1.0,
                help="Specific enthalpy"
            )

        with col2:
            pressure_preset = st.selectbox(
                "Pressure Preset",
                ["Custom"] + list(STEAM_PRESSURE_PRESETS.keys()),
                help="Select common pressure or enter custom"
            )

            if pressure_preset == "Custom":
                pressure_input = st.number_input(
                    "Pressure (psia)",
                    min_value=0.1,
                    max_value=3000.0,
                    value=14.7,
                    step=0.1
                )
            else:
                pressure_input = STEAM_PRESSURE_PRESETS[pressure_preset]
                st.info(f"Pressure: {pressure_input} psia")

        if st.button("Calculate from Enthalpy", type="primary", use_container_width=True):

            try:
                # Calculate quality
                quality_calc = steam_quality(enthalpy_input, pressure_input)

                # Get saturation temperature
                t_sat = saturation_temperature(pressure_input)

                # Determine phase and temperature
                if quality_calc < 0:
                    phase = "Subcooled Liquid"
                    phase_color = "#3498db"
                    # Estimate temperature (approximate)
                    temperature_est = t_sat - 10  # Simplified
                elif quality_calc > 1:
                    phase = "Superheated Vapor"
                    phase_color = "#e74c3c"
                    temperature_est = t_sat + 10  # Simplified
                else:
                    phase = "Saturated (Two-Phase)"
                    phase_color = "#f39c12"
                    temperature_est = t_sat

                # Display results
                st.success("Calculation complete")

                st.markdown(f"""
                <div style="background-color: {phase_color}; color: white;
                            padding: 1rem; border-radius: 10px; text-align: center;
                            font-size: 1.5rem; font-weight: bold; margin: 1rem 0;">
                    {phase}
                </div>
                """, unsafe_allow_html=True)

                st.subheader("Results")

                col1, col2, col3 = st.columns(3)

                with col1:
                    if -0.01 < quality_calc < 1.01:
                        st.metric(
                            "Quality",
                            f"{format_number(quality_calc * 100, 2)} %"
                        )
                    else:
                        st.metric(
                            "Quality",
                            f"{format_number(quality_calc, 3)}",
                            help=f"<0: Subcooled, >1: Superheated"
                        )

                with col2:
                    st.metric(
                        "Saturation Temp",
                        f"{format_number(t_sat, 1)} °F",
                        help=f"At {pressure_input} psia"
                    )

                with col3:
                    st.metric(
                        "Est. Temperature",
                        f"{format_number(temperature_est, 1)} °F",
                        help="Estimated (approximate for single-phase)"
                    )

            except Exception as e:
                st.error(f"Calculation Error: {str(e)}")
                st.exception(e)

    # Mode 3: Saturation properties only
    else:  # Saturation Properties Only

        st.subheader("📝 Input Parameter")

        mode = st.radio(
            "Calculate from:",
            ["Temperature (get Psat)", "Pressure (get Tsat)"],
            horizontal=True
        )

        if mode == "Temperature (get Psat)":
            temp_input = st.number_input(
                "Temperature (°F)",
                min_value=32.0,
                max_value=705.0,
                value=212.0,
                step=1.0
            )

            if st.button("Calculate Psat", type="primary", use_container_width=True):
                try:
                    p_sat = saturation_pressure(temp_input)

                    st.success("Calculation complete")
                    st.metric(
                        "Saturation Pressure",
                        f"{format_number(p_sat, 3)} psia",
                        help=f"At {temp_input}°F"
                    )

                    # Comparison if at 212°F
                    if abs(temp_input - 212.0) < 0.1:
                        st.markdown("---")
                        show_comparison_result(
                            p_sat, 14.696,
                            label="Psat",
                            unit="psia",
                            tolerance=0.005
                        )

                except Exception as e:
                    st.error(f"Error: {str(e)}")

        else:  # Pressure input
            press_input = st.number_input(
                "Pressure (psia)",
                min_value=0.1,
                max_value=3200.0,
                value=14.7,
                step=0.1
            )

            if st.button("Calculate Tsat", type="primary", use_container_width=True):
                try:
                    t_sat = saturation_temperature(press_input)

                    st.success("Calculation complete")
                    st.metric(
                        "Saturation Temperature",
                        f"{format_number(t_sat, 2)} °F",
                        help=f"At {press_input} psia"
                    )

                    # Comparison if at 14.7 psia
                    if abs(press_input - 14.7) < 0.1:
                        st.markdown("---")
                        show_comparison_result(
                            t_sat, 212.0,
                            label="Tsat",
                            unit="°F",
                            tolerance=0.005
                        )

                except Exception as e:
                    st.error(f"Error: {str(e)}")


if __name__ == "__main__":
    show_steam_properties_calculator()
