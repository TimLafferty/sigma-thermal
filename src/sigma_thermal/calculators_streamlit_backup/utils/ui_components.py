"""
UI Components for Sigma Thermal Calculators

Reusable Streamlit components for consistent UI across calculators.
"""

import streamlit as st
import pandas as pd
import plotly.graph_objects as go
import plotly.express as px
from typing import Dict, Any, List, Optional


def format_number(value: float, decimal_places: int = 2, use_commas: bool = True) -> str:
    """
    Format number for display.

    Parameters
    ----------
    value : float
        Number to format
    decimal_places : int
        Number of decimal places
    use_commas : bool
        Whether to use comma separators

    Returns
    -------
    str
        Formatted number string
    """
    if use_commas:
        return f"{value:,.{decimal_places}f}"
    else:
        return f"{value:.{decimal_places}f}"


def show_metric_card(label: str, value: float, unit: str = "", delta: Optional[float] = None,
                     decimal_places: int = 2):
    """
    Display a metric card with value and optional delta.

    Parameters
    ----------
    label : str
        Metric label
    value : float
        Metric value
    unit : str, optional
        Unit string (e.g., "BTU/lb")
    delta : float, optional
        Change value to display
    decimal_places : int
        Number of decimal places
    """
    formatted_value = format_number(value, decimal_places)

    if unit:
        formatted_value = f"{formatted_value} {unit}"

    st.metric(label, formatted_value, delta=delta)


def show_comparison_result(python_value: float, excel_value: float, label: str = "Result",
                           unit: str = "", tolerance: float = 0.01, decimal_places: int = 2):
    """
    Display Python vs Excel comparison with pass/fail indicator.

    Parameters
    ----------
    python_value : float
        Python calculated value
    excel_value : float
        Excel VBA value
    label : str
        Result label
    unit : str
        Unit string
    tolerance : float
        Acceptable relative tolerance (default 1%)
    decimal_places : int
        Decimal places for display
    """
    # Calculate deviation
    if excel_value != 0:
        deviation = abs(python_value - excel_value) / abs(excel_value)
    else:
        deviation = abs(python_value - excel_value)

    deviation_pct = deviation * 100

    # Determine status
    if deviation <= tolerance:
        status = "PASS"
        color = "#e8f5e9"  # subtle green
        border_color = "#4caf50"
    elif deviation <= tolerance * 2:
        status = "WARNING"
        color = "#fff8e1"  # subtle yellow
        border_color = "#ff9800"
    else:
        status = "FAIL"
        color = "#ffebee"  # subtle red
        border_color = "#f44336"

    # Create comparison table
    df = pd.DataFrame({
        "Source": ["Python", "Excel VBA", "Deviation"],
        f"{label} ({unit})" if unit else label: [
            format_number(python_value, decimal_places),
            format_number(excel_value, decimal_places),
            f"{deviation_pct:.4f}%"
        ]
    })

    st.dataframe(df, use_container_width=True, hide_index=True)

    # Status indicator
    st.markdown(f"""
    <div style="background-color: {color}; border-left: 4px solid {border_color};
                padding: 0.75rem 1rem; border-radius: 4px; margin-top: 0.5rem;">
        <strong>{status}</strong> - Deviation: {deviation_pct:.4f}% (Tolerance: {tolerance*100:.2f}%)
    </div>
    """, unsafe_allow_html=True)


def show_results_table(results: Dict[str, Any], units: Dict[str, str] = None,
                       decimal_places: int = 2):
    """
    Display results in a formatted table.

    Parameters
    ----------
    results : dict
        Dictionary of result name to value
    units : dict, optional
        Dictionary of result name to unit string
    decimal_places : int
        Decimal places
    """
    if units is None:
        units = {}

    data = []
    for key, value in results.items():
        unit = units.get(key, "")
        if isinstance(value, (int, float)):
            formatted_value = format_number(value, decimal_places)
            if unit:
                formatted_value = f"{formatted_value} {unit}"
        else:
            formatted_value = str(value)

        data.append({
            "Parameter": key,
            "Value": formatted_value
        })

    df = pd.DataFrame(data)
    st.table(df)


def show_input_validation(total: float, expected: float = 100.0, tolerance: float = 0.01,
                          label: str = "Total Composition"):
    """
    Show validation for summed inputs (e.g., fuel composition).

    Parameters
    ----------
    total : float
        Calculated total
    expected : float
        Expected total (default 100)
    tolerance : float
        Acceptable tolerance
    label : str
        Label for total
    """
    deviation = abs(total - expected)

    if deviation < tolerance:
        st.success(f"✅ {label} = {total:.2f}% (Valid)")
        return True
    else:
        st.error(f"⚠️ {label} = {total:.2f}% (Must equal {expected:.0f}%)")
        return False


def create_composition_pie_chart(composition: Dict[str, float], title: str = "Composition"):
    """
    Create pie chart for composition data.

    Parameters
    ----------
    composition : dict
        Component name to percentage
    title : str
        Chart title

    Returns
    -------
    plotly.graph_objects.Figure
        Pie chart figure
    """
    # Filter out zero values
    composition = {k: v for k, v in composition.items() if v > 0.01}

    labels = list(composition.keys())
    values = list(composition.values())

    fig = go.Figure(data=[go.Pie(
        labels=labels,
        values=values,
        hole=0.3,
        textinfo='label+percent',
        textposition='auto'
    )])

    fig.update_layout(
        title=title,
        showlegend=True,
        height=400
    )

    return fig


def create_bar_chart(data: Dict[str, float], title: str = "", x_label: str = "",
                     y_label: str = "", color: str = "#3498db"):
    """
    Create bar chart.

    Parameters
    ----------
    data : dict
        Category to value mapping
    title : str
        Chart title
    x_label : str
        X-axis label
    y_label : str
        Y-axis label
    color : str
        Bar color

    Returns
    -------
    plotly.graph_objects.Figure
        Bar chart figure
    """
    df = pd.DataFrame({
        'Category': list(data.keys()),
        'Value': list(data.values())
    })

    fig = px.bar(
        df,
        x='Category',
        y='Value',
        title=title,
        color_discrete_sequence=[color]
    )

    fig.update_layout(
        xaxis_title=x_label,
        yaxis_title=y_label,
        showlegend=False,
        height=400
    )

    return fig


def create_line_chart(x_data: List[float], y_data: List[float], title: str = "",
                      x_label: str = "", y_label: str = "", color: str = "#3498db"):
    """
    Create line chart.

    Parameters
    ----------
    x_data : list
        X-axis values
    y_data : list
        Y-axis values
    title : str
        Chart title
    x_label : str
        X-axis label
    y_label : str
        Y-axis label
    color : str
        Line color

    Returns
    -------
    plotly.graph_objects.Figure
        Line chart figure
    """
    fig = go.Figure()

    fig.add_trace(go.Scatter(
        x=x_data,
        y=y_data,
        mode='lines+markers',
        line=dict(color=color, width=3),
        marker=dict(size=8)
    ))

    fig.update_layout(
        title=title,
        xaxis_title=x_label,
        yaxis_title=y_label,
        showlegend=False,
        height=400
    )

    return fig


def export_results_json(results: Dict[str, Any], filename: str = "results.json"):
    """
    Create download button for results as JSON.

    Parameters
    ----------
    results : dict
        Results dictionary
    filename : str
        Download filename
    """
    import json

    json_str = json.dumps(results, indent=2)

    st.download_button(
        label="📥 Download Results (JSON)",
        data=json_str,
        file_name=filename,
        mime="application/json"
    )


def export_results_csv(results: Dict[str, Any], filename: str = "results.csv"):
    """
    Create download button for results as CSV.

    Parameters
    ----------
    results : dict
        Results dictionary
    filename : str
        Download filename
    """
    df = pd.DataFrame([results])

    csv_str = df.to_csv(index=False)

    st.download_button(
        label="📥 Download Results (CSV)",
        data=csv_str,
        file_name=filename,
        mime="text/csv"
    )


def show_info_box(message: str, box_type: str = "info"):
    """
    Display info/warning/error box.

    Parameters
    ----------
    message : str
        Message to display
    box_type : str
        Type: 'info', 'success', 'warning', 'error'
    """
    colors = {
        'info': ('#d1ecf1', '#0c5460'),
        'success': ('#d4edda', '#155724'),
        'warning': ('#fff3cd', '#856404'),
        'error': ('#f8d7da', '#721c24')
    }

    bg_color, text_color = colors.get(box_type, colors['info'])

    st.markdown(f"""
    <div style="background-color: {bg_color}; color: {text_color};
                padding: 1rem; border-radius: 5px; margin: 1rem 0;">
        {message}
    </div>
    """, unsafe_allow_html=True)


def show_equation(equation: str, label: str = ""):
    """
    Display formatted equation using LaTeX.

    Parameters
    ----------
    equation : str
        LaTeX equation string
    label : str
        Optional label
    """
    if label:
        st.markdown(f"**{label}:**")

    st.latex(equation)
