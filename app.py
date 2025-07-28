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

            st.plotly_chart(fig, use_container_width=True, key="description_chart")


            # New section for critical parts analysis across all projects
            st.markdown("---")
            st.subheader("🚨 Critical Parts Analysis (Frozen Zone) - All Projects")
            
            # Analyze frozen zone for all projects
            frozen_weeks = wk_cols[:4]
            # Use original df instead of df_selected to get all projects
            all_frozen_data = df.melt(
                id_vars=['Material', 'Material type', 'Deficit quantity'],
                value_vars=wk_cols,
                var_name='Week',
                value_name='Fluctuation'
            )
            
            # Filter for frozen weeks
            frozen_data = all_frozen_data[all_frozen_data['Week'].isin(frozen_weeks)].copy()
            
            # Convert fluctuation to numeric
            frozen_data['Fluctuation'] = pd.to_numeric(
                frozen_data['Fluctuation'].replace('%', '', regex=True),
                errors='coerce'
            ) / 100.0

            # Create critical parts dashboard
            critical_data = frozen_data[frozen_data['Fluctuation'] > 0.2].copy()
            
            if not critical_data.empty:
                # Create columns for metrics
                m1, m2, m3, m4 = st.columns(4)
                with m1:
                    st.metric("Critical Parts", len(critical_data['Material'].unique()))
                with m2:
                    st.metric("Projects Affected", len(critical_data['Material type'].unique()))
                with m3:
                    st.metric("Highest Fluctuation", f"{critical_data['Fluctuation'].max():.1%}")
                with m4:
                    st.metric("Affected Weeks", len(critical_data['Week'].unique()))

                # Create visualization columns
                col1, col2 = st.columns(2)
                
                with col1:
                    # Bar chart for critical parts
                    critical_summary = critical_data.groupby(['Material', 'Material type'])['Fluctuation'].max().reset_index()
                    fig_critical = px.bar(
                        critical_summary,
                        x='Material',
                        y='Fluctuation',
                        color='Material type',
                        title='Maximum Fluctuation by Critical Part',
                        labels={'Fluctuation': 'Max Fluctuation'},
                        height=400
                    )
                    fig_critical.add_hline(y=0.2, line_dash="dash", line_color="red")
                    st.plotly_chart(fig_critical, use_container_width=True, key="all_critical_chart")

                with col2:
                    # Heatmap by project and week
                    pivot_table = critical_data.pivot_table(
                        index=['Material type', 'Material'],
                        columns='Week',
                        values='Fluctuation',
                        aggfunc='max'
                    )
                    fig_heat = px.imshow(
                        pivot_table,
                        title='Critical Parts Heatmap by Project',
                        color_continuous_scale='RdYlBu_r',
                        aspect='auto',
                        labels={'color': 'Fluctuation'}
                    )
                    fig_heat.update_layout(height=400)
                    st.plotly_chart(fig_heat, use_container_width=True, key="all_heat_chart")

          # Detailed table with project information
            st.subheader("Critical Parts Details - All Projects")

            # Unique filter values
            material_types = critical_data['Material type'].unique()
            materials = critical_data['Material'].unique()
            weeks = sorted(critical_data['Week'].unique())

            # Filter widgets
            selected_material_type = st.multiselect("Select Project(s)", material_types)
            selected_material = st.multiselect("Select Part Number(s)", materials)
            selected_week = st.multiselect("Select Week(s)", weeks)

            # Apply filters
            filtered_data = critical_data.copy()

            if selected_material_type:
                filtered_data = filtered_data[filtered_data['Material type'].isin(selected_material_type)]

            if selected_material:
                filtered_data = filtered_data[filtered_data['Material'].isin(selected_material)]

            if selected_week:
                filtered_data = filtered_data[filtered_data['Week'].isin(selected_week)]

            # Prepare the table
            detail_data = filtered_data[
                ['Material type', 'Material', 'Week', 'Fluctuation']
            ].sort_values(['Material type', 'Material', 'Week'])

            detail_data['Fluctuation'] = detail_data['Fluctuation'].map('{:.1%}'.format)

            # Show the table
            st.dataframe(
                detail_data,
                column_config={
                    "Material type": "Project",
                    "Material": "Part Number",
                    "Week": "Critical Week",
                    "Fluctuation": "Fluctuation %"
                },
                hide_index=True
            )
