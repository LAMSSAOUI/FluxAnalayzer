import streamlit as st
import pandas as pd
import plotly.express as px
import io

st.set_page_config(page_title="Fluctuation Dashboard", layout="wide")
st.title("📈 Material Fluctuation Visualizer")

uploaded_file = st.file_uploader("Upload Excel File", type=["xlsx"])

if uploaded_file:
    xls = pd.ExcelFile(uploaded_file)
    sheet_name = st.selectbox("Select Sheet", xls.sheet_names)
    df = pd.read_excel(xls, sheet_name=sheet_name)

    if 'Material type' not in df.columns:
        st.error("❌ The selected sheet does not contain 'Material type' column.")
    else:
        project = st.selectbox("Select Project", sorted(df['Material type'].dropna().unique()))
        df_selected = df[df['Material type'] == project].copy()

        # Detect week columns
        wk_cols = [col for col in df_selected.columns if col.lower().startswith('wk')]
        if not wk_cols:
            st.warning("No week columns found (starting with 'wk').")
        else:
            # Clean % signs
            # for col in wk_cols + ['Deficit quantity']:
            #     df_selected[col] = df_selected[col].replace('%', '', regex=True).astype(float) / 100.0

            # Melt to long format
            df_long = df_selected.melt(
                id_vars=['Material', 'Material type', 'Deficit quantity'],
                value_vars=wk_cols,
                var_name='Week',
                value_name='Fluctuation'
            )

            # Add the "Deficit" as first week-like value
            deficit_rows = df_selected[['Material', 'Material type', 'Deficit quantity']].copy()
            deficit_rows['Week'] = 'Deficit'
            deficit_rows['Fluctuation'] = deficit_rows['Deficit quantity']
            df_long = pd.concat([deficit_rows, df_long], ignore_index=True)

            # Ensure proper ordering
            week_order = ['Deficit'] + wk_cols
            df_long['Week'] = pd.Categorical(df_long['Week'], categories=week_order, ordered=True)
            df_long = df_long.sort_values(['Material', 'Week'])

           

            fig = px.line(
                df_long,
                x='Week',
                y='Fluctuation',
                color='Material',
                markers=True,
                title=f"Fluctuation Over Weeks - {project}",
                hover_data=['Deficit quantity'],
                height=600,
                width=1100
            )

            # Add horizontal line at 20%
            shapes = [
                dict(
                    type='line',
                    xref='paper', x0=0, x1=1,
                    yref='y', y0=0.2, y1=0.2,
                    line=dict(color='red', width=2, dash='dash')
                ),
                dict(
                    type='line',
                    xref='paper', x0=0, x1=1,
                    yref='y', y0=-0.2, y1=-0.2,
                    line=dict(color='red', width=2, dash='dash')
                )
            ]

            # Add frozen zone (first 4 weeks)
            if len(wk_cols) >= 4:
                shapes.append(
                    dict(
                        type='rect',
                        xref='x',
                        yref='paper',
                        x0=wk_cols[0],
                        x1=wk_cols[3],
                        y0=0,
                        y1=1,
                        fillcolor='rgba(255, 0, 0, 0.1)',
                        line=dict(width=0),
                        layer='below'
                    )
                )

            fig.update_layout(
                yaxis=dict(tickformat=".0%"),
                xaxis_tickangle=45,
                shapes=shapes
            )

            st.plotly_chart(fig, use_container_width=True)
