import sys
import pandas as pd
import plotly.express as px
from PySide6.QtWidgets import (
    QApplication, QWidget, QVBoxLayout, QPushButton, QFileDialog, QLabel,
    QComboBox, QTableView, QMessageBox, QHBoxLayout, QListWidget, QListWidgetItem,
    QAbstractItemView
)
from PySide6.QtCore import Qt, QAbstractTableModel, QUrl
from PySide6.QtWebEngineWidgets import QWebEngineView


class PandasModel(QAbstractTableModel):
    def __init__(self, df=pd.DataFrame(), parent=None):
        super().__init__(parent)
        self._df = df

    def rowCount(self, parent=None):
        return len(self._df.index)

    def columnCount(self, parent=None):
        return len(self._df.columns)

    def data(self, index, role=Qt.DisplayRole):
        if index.isValid() and role == Qt.DisplayRole:
            return str(self._df.iat[index.row(), index.column()])
        return None

    def headerData(self, section, orientation, role=Qt.DisplayRole):
        if role == Qt.DisplayRole:
            if orientation == Qt.Horizontal:
                return str(self._df.columns[section])
            else:
                return str(self._df.index[section])
        return None


class FluctuationApp(QWidget):
    def __init__(self):
        super().__init__()
        self.setWindowTitle("Material Fluctuation Visualizer")
        self.resize(1200, 900)

        self.layout = QVBoxLayout(self)


        self.load_btn = QPushButton("Upload Excel File")
        self.load_btn.clicked.connect(self.load_file)
        self.layout.addWidget(self.load_btn)


        self.sheet_combo = QComboBox()
        # self.sheet_combo.currentIndexChanged.connect(self.load_sheet_data)
        self.layout.addWidget(QLabel("Select Sheet:"))
        self.layout.addWidget(self.sheet_combo)
        # self.sheet_combo.currentIndexChanged.connect(self.load_sheet_data)
        self.sheet_combo.activated.connect(self.load_sheet_data)
        # Project selector
        self.project_combo = QComboBox()
        self.project_combo.currentIndexChanged.connect(self.update_project_selection)
        self.layout.addWidget(QLabel("Select Project:"))
        self.layout.addWidget(self.project_combo)

      
        self.chart_view = QWebEngineView()
        self.layout.addWidget(self.chart_view, stretch=3)

   
        self.summary_layout = QHBoxLayout()
        self.metric_critical_parts = QLabel("Critical Parts: 0")
        self.metric_projects_affected = QLabel("Projects Affected: 0")
        self.metric_highest_fluctuation = QLabel("Highest Fluctuation: 0%")
        self.metric_affected_weeks = QLabel("Affected Weeks: 0")

        for w in [self.metric_critical_parts, self.metric_projects_affected,
                  self.metric_highest_fluctuation, self.metric_affected_weeks]:
            w.setStyleSheet("font-weight: bold; font-size: 14px; padding: 5px;")
            self.summary_layout.addWidget(w)
        self.layout.addLayout(self.summary_layout)

        
        self.critical_table = QTableView()
        self.layout.addWidget(QLabel("Critical Parts Details"))
        self.layout.addWidget(self.critical_table, stretch=2)

     
        self.df = None
        self.current_sheet_df = None

    def load_file(self):
        file_path, _ = QFileDialog.getOpenFileName(self, "Open Excel File", filter="Excel Files (*.xlsx *.xlsm)")
        if not file_path:
            return

        try:
            self.xls = pd.ExcelFile(file_path)
            self.sheet_combo.clear()
            self.sheet_combo.addItems(self.xls.sheet_names)
            # Auto-select first sheet
            # self.sheet_combo.setCurrentIndex(0)
            # self.load_sheet_data()
        except Exception as e:
            QMessageBox.warning(self, "Error", f"Failed to load Excel file:\n{e}")

    def load_sheet_data(self):
        if not self.xls or self.sheet_combo.currentIndex() < 0:
            return
        sheet = self.sheet_combo.currentText()
        try:
            df = pd.read_excel(self.xls, sheet_name=sheet)
            self.df = df
            if 'Production line' not in df.columns:
                QMessageBox.warning(self, "Error", "Sheet missing 'Production line' column.")
                self.project_combo.clear()
                return

            projects = sorted(df['Production line'].dropna().unique())
            self.project_combo.clear()
            self.project_combo.addItems(projects)
            self.project_combo.setCurrentIndex(0)
            self.update_project_selection()
        except Exception as e:
            QMessageBox.warning(self, "Error", f"Failed to load sheet data:\n{e}")

    def update_project_selection(self):
        if self.df is None:
            return
        project = self.project_combo.currentText()
        if not project:
            return

        df_selected = self.df[self.df['Production line'] == project].copy()
        wk_cols = [col for col in df_selected.columns if col.lower().startswith('wk')]
        if not wk_cols:
            QMessageBox.information(self, "Info", "No week columns found starting with 'wk'.")
            return


        df_long = df_selected.melt(
            id_vars=['Material', 'Production line', 'Deficit quantity'],
            value_vars=wk_cols,
            var_name='Week',
            value_name='Fluctuation'
        )
        deficit_rows = df_selected[['Material', 'Production line', 'Deficit quantity']].copy()
        deficit_rows['Week'] = 'Deficit'
        deficit_rows['Fluctuation'] = deficit_rows['Deficit quantity']
        df_long = pd.concat([deficit_rows, df_long], ignore_index=True)

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

        fig.add_hline(y=0.2, line_dash="dash", line_color="red")
        fig.add_hline(y=-0.2, line_dash="dash", line_color="red")


        if len(wk_cols) >= 4:
            fig.add_vrect(
                x0=wk_cols[0], x1=wk_cols[3],
                fillcolor="red", opacity=0.1, layer="below", line_width=0
            )

        fig.update_layout(
            yaxis=dict(tickformat=".0%"),
            xaxis_tickangle=45
        )

        # Render plotly figure to html and load in QWebEngineView
        html = fig.to_html(include_plotlyjs='cdn')
        self.chart_view.setHtml(html)

        # Show critical parts summary (simplified: parts where fluctuation > 0.2 in frozen zone)
        frozen_weeks = wk_cols[:4]
        all_frozen_data = self.df.melt(
            id_vars=['Material', 'Production line', 'Deficit quantity'],
            value_vars=wk_cols,
            var_name='Week',
            value_name='Fluctuation'
        )
        frozen_data = all_frozen_data[all_frozen_data['Week'].isin(frozen_weeks)].copy()

        critical_data = frozen_data[frozen_data['Fluctuation'] > 0.2].copy()

        self.metric_critical_parts.setText(f"Critical Parts: {len(critical_data['Material'].unique())}")
        self.metric_projects_affected.setText(f"Projects Affected: {len(critical_data['Production line'].unique())}")
        highest = critical_data['Fluctuation'].max() if not critical_data.empty else 0
        self.metric_highest_fluctuation.setText(f"Highest Fluctuation: {highest:.1%}")
        self.metric_affected_weeks.setText(f"Affected Weeks: {len(critical_data['Week'].unique())}")

        # Show critical parts table
        if not critical_data.empty:
            display_df = critical_data[['Production line', 'Material', 'Week', 'Fluctuation']].copy()
            display_df['Fluctuation'] = display_df['Fluctuation'].map('{:.1%}'.format)
            self.critical_table.setModel(PandasModel(display_df))
        else:
            self.critical_table.setModel(PandasModel(pd.DataFrame()))

if __name__ == "__main__":
    app = QApplication(sys.argv)
    window = FluctuationApp()
    window.show()
    sys.exit(app.exec())
