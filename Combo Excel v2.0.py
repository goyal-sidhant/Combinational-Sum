"""
Combination Sum Finder - Excel Master Edition v2.1
FIXED: Text visibility and proper filtered data reading
"""

import sys
import time
from typing import List, Tuple
from dataclasses import dataclass

from PyQt6.QtWidgets import *
from PyQt6.QtCore import *
from PyQt6.QtGui import *

try:
    import xlwings as xw
    XLWINGS_AVAILABLE = True
except ImportError:
    XLWINGS_AVAILABLE = False
    print("Warning: xlwings not available. Excel features will be disabled.")

from collections import Counter


@dataclass
class Combination:
    """Represents a single combination result"""
    numbers: List[float]
    sum_value: float
    indices: List[int]  # Maps to Excel cell positions
    is_exact: bool


class ExcelBridge(QObject):
    """Handles all Excel interactions"""
    data_loaded = pyqtSignal(list, list)  # numbers, cell_addresses
    error_occurred = pyqtSignal(str)

    def __init__(self):
        super().__init__()
        self.app = None
        self.book = None
        self.sheet = None
        self.original_range = None
        self.cell_addresses = []  # Store actual Excel addresses

    def connect_to_excel(self):
        """Connect to active Excel instance"""
        if not XLWINGS_AVAILABLE:
            self.error_occurred.emit("xlwings is not installed. Please install with: pip install xlwings")
            return False

        try:
            # Try to get active Excel app
            self.app = xw.apps.active
            if not self.app:
                # No active Excel, try to create new instance
                self.app = xw.App(visible=True)
            return True
        except Exception as e:
            self.error_occurred.emit(f"Could not connect to Excel: {str(e)}\n\nMake sure Excel is running.")
            return False

    def get_open_workbooks(self):
        """Get list of open workbook names"""
        try:
            if not self.app:
                self.connect_to_excel()
            return [book.name for book in self.app.books]
        except:
            return []

    def select_workbook(self, workbook_name):
        """Select a specific workbook"""
        try:
            self.book = self.app.books[workbook_name]
            return True
        except Exception as e:
            self.error_occurred.emit(f"Could not select workbook: {str(e)}")
            return False

    def get_sheets(self):
        """Get list of sheet names in current workbook"""
        if not self.book:
            return []
        return [sheet.name for sheet in self.book.sheets]

    def read_selection(self, sheet_name, selection_range=None, filtered_only=True):
        """Read data from Excel selection or specified range - PROPERLY handling filters"""
        try:
            self.sheet = self.book.sheets[sheet_name]

            # Get range
            if selection_range:
                rng = self.sheet.range(selection_range)
            else:
                # Get current selection from the sheet's application
                selection = self.sheet.api.Application.Selection
                rng = self.sheet.range(selection.Address)

            self.original_range = rng

            values = []
            cell_addresses = []

            # If filtered_only is True and AutoFilter is active
            if filtered_only and hasattr(self.sheet.api, 'AutoFilterMode') and self.sheet.api.AutoFilterMode:
                # Use xlwings built-in method to get visible cells only
                # This is the correct way to handle filtered data
                try:
                    # Get each cell in the range and check if it's visible
                    for cell in rng:
                        # Check if the row is hidden (filtered out)
                        if not cell.api.EntireRow.Hidden:
                            val = cell.value
                            if val is not None and self._is_number(val):
                                values.append(float(val))
                                # Calculate relative position
                                row_offset = cell.row - rng.row
                                col_offset = cell.column - rng.column
                                cell_addresses.append((row_offset, col_offset))
                except Exception as e:
                    # Fallback method using COM API directly
                    visible_cells = rng.api.SpecialCells(12)  # xlCellTypeVisible
                    for cell in visible_cells:
                        val = cell.Value
                        if val is not None and self._is_number(val):
                            values.append(float(val))
                            row_offset = cell.Row - rng.api.Cells(1, 1).Row
                            col_offset = cell.Column - rng.api.Cells(1, 1).Column
                            cell_addresses.append((row_offset, col_offset))
            else:
                # No filter or filtered_only is False - read all cells normally
                all_values = rng.value

                if all_values is None:
                    self.error_occurred.emit("No data found in selection")
                    return []

                # Handle different data structures (single value, 1D array, 2D array)
                if not isinstance(all_values, list):
                    if self._is_number(all_values):
                        values = [float(all_values)]
                        cell_addresses = [(0, 0)]
                else:
                    # Handle 1D array
                    if not any(isinstance(item, list) for item in all_values):
                        for i, val in enumerate(all_values):
                            if val is not None and self._is_number(val):
                                values.append(float(val))
                                cell_addresses.append((i, 0))
                    else:
                        # Handle 2D array
                        for i, row in enumerate(all_values):
                            if not isinstance(row, list):
                                row = [row]
                            for j, val in enumerate(row):
                                if val is not None and self._is_number(val):
                                    values.append(float(val))
                                    cell_addresses.append((i, j))

            self.cell_addresses = cell_addresses
            self.data_loaded.emit(values, cell_addresses)
            return values

        except Exception as e:
            self.error_occurred.emit(f"Error reading Excel: {str(e)}")
            return []

    def _is_number(self, value):
        """Check if value is a number"""
        try:
            float(value)
            return True
        except:
            return False

    def highlight_cells(self, indices: List[int], color_hex='#FFFF00'):
        """Temporarily highlight specific cells in Excel"""
        try:
            if not self.original_range or not self.cell_addresses:
                return

            # Clear previous highlighting first
            self.clear_highlighting()

            # Convert hex color to RGB
            rgb = self._hex_to_rgb(color_hex)

            # Highlight selected cells
            for idx in indices:
                if idx < len(self.cell_addresses):
                    row_offset, col_offset = self.cell_addresses[idx]
                    cell = self.original_range[row_offset, col_offset]
                    cell.color = rgb

        except Exception as e:
            self.error_occurred.emit(f"Error highlighting cells: {str(e)}")

    def color_cells_permanent(self, indices: List[int], color_hex='#90EE90'):
        """Permanently color cells (for marking as 'used')"""
        try:
            if not self.original_range or not self.cell_addresses:
                return

            # Convert hex color to RGB
            rgb = self._hex_to_rgb(color_hex)

            for idx in indices:
                if idx < len(self.cell_addresses):
                    row_offset, col_offset = self.cell_addresses[idx]
                    cell = self.original_range[row_offset, col_offset]
                    cell.color = rgb

        except Exception as e:
            self.error_occurred.emit(f"Error coloring cells: {str(e)}")

    def clear_highlighting(self):
        """Clear all cell colors in the original range"""
        try:
            if self.original_range:
                self.original_range.color = None  # Remove all colors
        except:
            pass

    def _hex_to_rgb(self, hex_color):
        """Convert hex color to RGB tuple for xlwings"""
        hex_color = hex_color.lstrip('#')
        return tuple(int(hex_color[i:i+2], 16) for i in (0, 2, 4))


class CombinationFinder(QThread):
    """Thread-based combination finder with progress reporting"""
    progress = pyqtSignal(int, float)  # found count, elapsed time
    result_found = pyqtSignal(Combination)
    finished = pyqtSignal(list)

    def __init__(self):
        super().__init__()
        self.numbers = []
        self.indices = []
        self.target = 0
        self.tolerance = 0
        self.max_length = 15
        self.max_results = 1000
        self.stop_flag = False
        self.results = []

    def setup(self, numbers, indices, target, tolerance, max_length=15, max_results=1000):
        self.numbers = numbers
        self.indices = indices
        self.target = target
        self.tolerance = tolerance
        self.max_length = max_length
        self.max_results = max_results
        self.stop_flag = False
        self.results = []

    def stop(self):
        self.stop_flag = True

    def run(self):
        """Main search algorithm"""
        start_time = time.time()

        # Sort numbers descending for shorter combinations first
        sorted_data = sorted(zip(self.numbers, self.indices), reverse=True)
        sorted_numbers = [x[0] for x in sorted_data]
        sorted_indices = [x[1] for x in sorted_data]

        def find_recursive(idx, current_sum, current_combo, current_indices):
            if self.stop_flag or len(self.results) >= self.max_results:
                return

            # Check if valid combination
            if current_combo and abs(current_sum - self.target) <= self.tolerance:
                is_exact = abs(current_sum - self.target) < 0.01
                combo = Combination(
                    numbers=current_combo.copy(),
                    sum_value=current_sum,
                    indices=current_indices.copy(),
                    is_exact=is_exact
                )
                self.results.append(combo)
                self.result_found.emit(combo)

                elapsed = time.time() - start_time
                self.progress.emit(len(self.results), elapsed)

                if len(self.results) >= self.max_results:
                    return

            # Pruning conditions
            if len(current_combo) >= self.max_length:
                return

            if idx >= len(sorted_numbers):
                return

            # Try remaining numbers
            for i in range(idx, len(sorted_numbers)):
                if self.stop_flag:
                    return

                new_sum = current_sum + sorted_numbers[i]

                # Skip if too large
                if new_sum > self.target + self.tolerance:
                    continue

                # Skip duplicates at same level
                if i > idx and abs(sorted_numbers[i] - sorted_numbers[i - 1]) < 1e-9:
                    continue

                current_combo.append(sorted_numbers[i])
                current_indices.append(sorted_indices[i])

                find_recursive(i + 1, new_sum, current_combo, current_indices)

                current_combo.pop()
                current_indices.pop()

        find_recursive(0, 0, [], [])

        # Sort results by length, then by exactness
        self.results.sort(key=lambda x: (len(x.numbers), not x.is_exact))

        self.finished.emit(self.results)


class MainWindow(QMainWindow):
    def __init__(self):
        super().__init__()
        self.excel_bridge = ExcelBridge()
        self.finder_thread = None
        self.current_numbers = []
        self.current_indices = []
        self.results = []
        self.current_selected_indices = []

        self.init_ui()
        self.setup_connections()
        self.apply_dark_theme()

    def apply_dark_theme(self):
        """Apply a dark theme with better contrast"""
        dark_style = """
        QMainWindow {
            background-color: #2b2b2b;
        }
        QGroupBox {
            background-color: #353535;
            border: 2px solid #555;
            border-radius: 5px;
            margin-top: 10px;
            padding-top: 10px;
            color: #ffffff;
            font-weight: bold;
        }
        QGroupBox::title {
            subcontrol-origin: margin;
            left: 10px;
            padding: 0 5px 0 5px;
        }
        QLabel {
            color: #ffffff;
        }
        QPushButton {
            background-color: #454545;
            color: #ffffff;
            border: 1px solid #555;
            padding: 5px;
            border-radius: 3px;
        }
        QPushButton:hover {
            background-color: #555555;
        }
        QPushButton:pressed {
            background-color: #353535;
        }
        QPushButton:disabled {
            background-color: #2b2b2b;
            color: #888888;
        }
        QComboBox, QLineEdit, QSpinBox, QDoubleSpinBox {
            background-color: #454545;
            color: #ffffff;
            border: 1px solid #555;
            padding: 3px;
            border-radius: 3px;
        }
        QTextEdit {
            background-color: #454545;
            color: #ffffff;
            border: 1px solid #555;
        }
        QListWidget {
            background-color: #353535;
            color: #ffffff;
            border: 1px solid #555;
            outline: none;
        }
        QListWidget::item {
            border-bottom: 1px solid #454545;
            padding: 5px;
        }
        QListWidget::item:selected {
            background-color: #555555;
        }
        QProgressBar {
            background-color: #454545;
            border: 1px solid #555;
            border-radius: 3px;
            text-align: center;
            color: white;
        }
        QProgressBar::chunk {
            background-color: #4CAF50;
            border-radius: 3px;
        }
        QCheckBox {
            color: #ffffff;
        }
        """
        self.setStyleSheet(dark_style)

    def init_ui(self):
        self.setWindowTitle("Combination Sum Finder - Excel Master Edition v2.1")
        self.setGeometry(100, 100, 1300, 850)

        # Central widget with splitter
        central_widget = QWidget()
        self.setCentralWidget(central_widget)
        main_layout = QHBoxLayout(central_widget)

        splitter = QSplitter(Qt.Orientation.Horizontal)
        main_layout.addWidget(splitter)

        # Left Panel - Excel Connection
        left_panel = self.create_excel_panel()
        splitter.addWidget(left_panel)

        # Right Panel - Search and Results
        right_panel = self.create_search_panel()
        splitter.addWidget(right_panel)

        # Set splitter sizes (30% left, 70% right)
        splitter.setSizes([400, 900])

        # Status Bar
        self.status_bar = self.statusBar()
        self.status_bar.setStyleSheet("color: #ffffff; background-color: #2b2b2b;")
        self.status_bar.showMessage("Ready. Connect to Excel to start.")

    def create_excel_panel(self):
        """Create Excel connection panel"""
        panel = QGroupBox("Excel Connection")
        layout = QVBoxLayout()

        # Connection status
        status_layout = QHBoxLayout()
        status_layout.addWidget(QLabel("Status:"))
        self.lbl_excel_status = QLabel("Not connected")
        self.lbl_excel_status.setStyleSheet("color: #ff6b6b; font-weight: bold;")
        status_layout.addWidget(self.lbl_excel_status)
        status_layout.addStretch()
        layout.addLayout(status_layout)

        # Connect button
        self.btn_connect = QPushButton("Connect to Excel")
        self.btn_connect.setMinimumHeight(35)
        self.btn_connect.clicked.connect(self.connect_to_excel)
        layout.addWidget(self.btn_connect)

        layout.addWidget(QLabel(""))  # Spacer

        # Workbook selector
        layout.addWidget(QLabel("Workbook:"))
        self.combo_workbook = QComboBox()
        self.combo_workbook.currentTextChanged.connect(self.on_workbook_changed)
        layout.addWidget(self.combo_workbook)

        # Sheet selector
        layout.addWidget(QLabel("Sheet:"))
        self.combo_sheet = QComboBox()
        layout.addWidget(self.combo_sheet)

        # Range input
        layout.addWidget(QLabel("Range (optional):"))
        self.txt_range = QLineEdit()
        self.txt_range.setPlaceholderText("e.g., A1:A100 or leave empty")
        layout.addWidget(self.txt_range)

        hint_label = QLabel("Leave empty to use current\nExcel selection")
        hint_label.setStyleSheet("color: #888888; font-size: 9pt;")
        layout.addWidget(hint_label)

        # Filtered checkbox
        self.chk_filtered = QCheckBox("Read filtered cells only")
        self.chk_filtered.setChecked(True)  # Default to true
        layout.addWidget(self.chk_filtered)

        # Import button
        self.btn_import = QPushButton("Import Data from Excel")
        self.btn_import.setMinimumHeight(35)
        self.btn_import.setStyleSheet("background-color: #4CAF50; font-weight: bold;")
        self.btn_import.clicked.connect(self.import_data)
        self.btn_import.setEnabled(False)
        layout.addWidget(self.btn_import)

        layout.addWidget(QLabel(""))  # Spacer

        # Data preview
        preview_group = QGroupBox("Data Preview")
        preview_layout = QVBoxLayout()
        self.txt_preview = QTextEdit()
        self.txt_preview.setMaximumHeight(120)
        self.txt_preview.setReadOnly(True)
        self.txt_preview.setPlaceholderText("Imported numbers will appear here...")
        preview_layout.addWidget(self.txt_preview)
        preview_group.setLayout(preview_layout)
        layout.addWidget(preview_group)

        # Color controls
        color_group = QGroupBox("Cell Colors")
        color_layout = QVBoxLayout()

        # Highlight color
        highlight_layout = QHBoxLayout()
        highlight_layout.addWidget(QLabel("Highlight:"))
        self.btn_highlight_color = QPushButton("   ")
        self.btn_highlight_color.setStyleSheet("background-color: #FFFF00; border: 1px solid black;")
        self.btn_highlight_color.setMaximumWidth(40)
        self.btn_highlight_color.clicked.connect(lambda: self.pick_color('highlight'))
        self.highlight_color = '#FFFF00'
        highlight_layout.addWidget(self.btn_highlight_color)
        highlight_layout.addStretch()
        color_layout.addLayout(highlight_layout)

        # Permanent color
        permanent_layout = QHBoxLayout()
        permanent_layout.addWidget(QLabel("Mark Used:"))
        self.btn_permanent_color = QPushButton("   ")
        self.btn_permanent_color.setStyleSheet("background-color: #90EE90; border: 1px solid black;")
        self.btn_permanent_color.setMaximumWidth(40)
        self.btn_permanent_color.clicked.connect(lambda: self.pick_color('permanent'))
        self.permanent_color = '#90EE90'
        permanent_layout.addWidget(self.btn_permanent_color)
        permanent_layout.addStretch()
        color_layout.addLayout(permanent_layout)

        self.btn_mark_used = QPushButton("Mark Selected as Used")
        self.btn_mark_used.setMinimumHeight(30)
        self.btn_mark_used.clicked.connect(self.mark_cells_as_used)
        self.btn_mark_used.setEnabled(False)
        color_layout.addWidget(self.btn_mark_used)

        self.btn_clear_colors = QPushButton("Clear All Colors")
        self.btn_clear_colors.clicked.connect(self.clear_colors)
        color_layout.addWidget(self.btn_clear_colors)

        color_group.setLayout(color_layout)
        layout.addWidget(color_group)

        layout.addStretch()
        panel.setLayout(layout)
        return panel

    def create_search_panel(self):
        """Create search parameters and results panel"""
        panel = QWidget()
        layout = QVBoxLayout()

        # Search parameters
        param_group = QGroupBox("Search Parameters")
        param_layout = QGridLayout()

        param_layout.addWidget(QLabel("Target Sum:"), 0, 0)
        self.spin_target = QDoubleSpinBox()
        self.spin_target.setRange(-999999, 999999)
        self.spin_target.setDecimals(2)
        self.spin_target.setMinimumWidth(120)
        param_layout.addWidget(self.spin_target, 0, 1)

        param_layout.addWidget(QLabel("Tolerance:"), 0, 2)
        self.spin_tolerance = QDoubleSpinBox()
        self.spin_tolerance.setRange(0, 9999)
        self.spin_tolerance.setDecimals(2)
        self.spin_tolerance.setValue(0)
        self.spin_tolerance.setMinimumWidth(100)
        param_layout.addWidget(self.spin_tolerance, 0, 3)

        param_layout.addWidget(QLabel("Max Length:"), 1, 0)
        self.spin_max_length = QSpinBox()
        self.spin_max_length.setRange(1, 100)
        self.spin_max_length.setValue(15)
        param_layout.addWidget(self.spin_max_length, 1, 1)

        param_layout.addWidget(QLabel("Max Results:"), 1, 2)
        self.spin_max_results = QSpinBox()
        self.spin_max_results.setRange(1, 10000)
        self.spin_max_results.setValue(100)
        param_layout.addWidget(self.spin_max_results, 1, 3)

        param_group.setLayout(param_layout)
        layout.addWidget(param_group)

        # Search buttons
        button_layout = QHBoxLayout()

        self.btn_search = QPushButton("Find Combinations")
        self.btn_search.setMinimumHeight(40)
        self.btn_search.setStyleSheet("font-size: 12pt; font-weight: bold; background-color: #4CAF50;")
        self.btn_search.clicked.connect(self.start_search)
        self.btn_search.setEnabled(False)
        button_layout.addWidget(self.btn_search)

        self.btn_stop = QPushButton("Stop")
        self.btn_stop.setMinimumHeight(40)
        self.btn_stop.setStyleSheet("background-color: #f44336;")
        self.btn_stop.clicked.connect(self.stop_search)
        self.btn_stop.setEnabled(False)
        button_layout.addWidget(self.btn_stop)

        layout.addLayout(button_layout)

        # Progress
        self.lbl_progress = QLabel("Ready to search...")
        layout.addWidget(self.lbl_progress)

        self.progress_bar = QProgressBar()
        layout.addWidget(self.progress_bar)

        # Results
        results_group = QGroupBox("Results - Click any row to highlight in Excel")
        results_layout = QVBoxLayout()

        # Filter buttons
        filter_layout = QHBoxLayout()
        self.btn_show_all = QPushButton("All")
        self.btn_show_exact = QPushButton("Exact Only")
        self.btn_show_approx = QPushButton("Approximate Only")

        self.btn_show_all.clicked.connect(lambda: self.filter_results('all'))
        self.btn_show_exact.clicked.connect(lambda: self.filter_results('exact'))
        self.btn_show_approx.clicked.connect(lambda: self.filter_results('approx'))

        filter_layout.addWidget(self.btn_show_all)
        filter_layout.addWidget(self.btn_show_exact)
        filter_layout.addWidget(self.btn_show_approx)
        filter_layout.addStretch()
        results_layout.addLayout(filter_layout)

        # Results list with better visibility
        self.results_list = QListWidget()
        self.results_list.setAlternatingRowColors(True)
        self.results_list.itemClicked.connect(self.on_result_selected)
        results_layout.addWidget(self.results_list)

        results_group.setLayout(results_layout)
        layout.addWidget(results_group)

        panel.setLayout(layout)
        return panel

    def setup_connections(self):
        """Connect signals"""
        self.excel_bridge.data_loaded.connect(self.on_data_loaded)
        self.excel_bridge.error_occurred.connect(self.on_error)

    def connect_to_excel(self):
        """Connect to Excel application"""
        if self.excel_bridge.connect_to_excel():
            self.lbl_excel_status.setText("Connected")
            self.lbl_excel_status.setStyleSheet("color: #4CAF50; font-weight: bold;")

            # Load workbooks
            workbooks = self.excel_bridge.get_open_workbooks()
            self.combo_workbook.clear()
            self.combo_workbook.addItems(workbooks)

            self.btn_import.setEnabled(True)
            self.status_bar.showMessage("Connected to Excel successfully!")
        else:
            self.lbl_excel_status.setText("Failed")
            self.lbl_excel_status.setStyleSheet("color: #ff6b6b; font-weight: bold;")

    def on_workbook_changed(self, workbook_name):
        """When user selects a workbook"""
        if workbook_name and self.excel_bridge.select_workbook(workbook_name):
            sheets = self.excel_bridge.get_sheets()
            self.combo_sheet.clear()
            self.combo_sheet.addItems(sheets)

    def import_data(self):
        """Import data from Excel"""
        sheet_name = self.combo_sheet.currentText()
        if not sheet_name:
            QMessageBox.warning(self, "Warning", "Please select a sheet first")
            return

        range_text = self.txt_range.text().strip()
        filtered_only = self.chk_filtered.isChecked()

        self.status_bar.showMessage(f"Importing data (filtered={filtered_only})...")
        self.excel_bridge.read_selection(sheet_name, range_text or None, filtered_only)

    def on_data_loaded(self, numbers, cell_addresses):
        """Data successfully loaded from Excel"""
        self.current_numbers = numbers
        self.current_indices = list(range(len(numbers)))  # Simple indices

        # Show preview
        preview_text = f"Loaded {len(numbers)} numbers\n\n"
        if len(numbers) <= 30:
            shown = numbers
        else:
            shown = numbers[:30]
            preview_text += "(Showing first 30)\n"

        preview_text += ", ".join(f"{n:.2f}" if n != int(n) else str(int(n)) for n in shown)
        if len(numbers) > 30:
            preview_text += f"\n... and {len(numbers)-30} more"

        self.txt_preview.setText(preview_text)
        self.btn_search.setEnabled(True)

        filter_status = "filtered cells only" if self.chk_filtered.isChecked() else "all cells"
        self.status_bar.showMessage(f"Loaded {len(numbers)} numbers from Excel ({filter_status})")

    def on_error(self, error_msg):
        """Handle Excel errors"""
        QMessageBox.critical(self, "Error", error_msg)
        self.status_bar.showMessage("Error occurred")

    def start_search(self):
        """Start searching for combinations"""
        if not self.current_numbers:
            QMessageBox.warning(self, "Warning", "Please import data from Excel first")
            return

        # Get parameters
        target = self.spin_target.value()
        tolerance = self.spin_tolerance.value()
        max_length = self.spin_max_length.value()
        max_results = self.spin_max_results.value()

        # Setup thread
        self.finder_thread = CombinationFinder()
        self.finder_thread.progress.connect(self.on_search_progress)
        self.finder_thread.result_found.connect(self.on_result_found)
        self.finder_thread.finished.connect(self.on_search_finished)

        self.finder_thread.setup(
            self.current_numbers,
            self.current_indices,
            target,
            tolerance,
            max_length,
            max_results
        )

        # Clear previous results
        self.results_list.clear()
        self.results = []

        # Update UI
        self.btn_search.setEnabled(False)
        self.btn_stop.setEnabled(True)
        self.btn_import.setEnabled(False)
        self.progress_bar.setMaximum(max_results)
        self.progress_bar.setValue(0)

        if tolerance > 0:
            self.lbl_progress.setText(f"Searching for {target} +/- {tolerance}...")
        else:
            self.lbl_progress.setText(f"Searching for exact match: {target}...")

        # Start search
        self.finder_thread.start()

    def stop_search(self):
        """Stop the search"""
        if self.finder_thread:
            self.finder_thread.stop()
            self.lbl_progress.setText("Stopping search...")

    def on_search_progress(self, found, elapsed):
        """Update progress during search"""
        self.progress_bar.setValue(found)
        rate = found / elapsed if elapsed > 0 else 0
        self.lbl_progress.setText(f"Found {found} combinations in {elapsed:.1f}s ({rate:.1f}/sec)")

    def on_result_found(self, combination):
        """New result found - with better text visibility"""
        self.results.append(combination)

        # Format display
        numbers_str = "{" + ", ".join(f"{n:.2f}" if n != int(n) else str(int(n))
                                      for n in combination.numbers) + "}"

        if combination.is_exact:
            # WHITE text on dark green background for better contrast
            item_text = f"[{len(combination.numbers)}] {numbers_str} = {combination.sum_value:.2f}"
            item = QListWidgetItem(item_text)
            item.setBackground(QColor(0, 100, 0))  # Dark green background
            item.setForeground(QColor(255, 255, 255))  # WHITE text
        else:
            # WHITE text on dark orange background for better contrast
            diff = combination.sum_value - self.spin_target.value()
            sign = "+" if diff > 0 else ""
            item_text = f"[{len(combination.numbers)}] {numbers_str} = {combination.sum_value:.2f} ({sign}{diff:.2f})"
            item = QListWidgetItem(item_text)
            item.setBackground(QColor(139, 69, 19))  # Dark orange/brown background
            item.setForeground(QColor(255, 255, 255))  # WHITE text

        item.setData(Qt.ItemDataRole.UserRole, len(self.results) - 1)
        self.results_list.addItem(item)

    def on_search_finished(self, results):
        """Search completed"""
        self.btn_search.setEnabled(True)
        self.btn_stop.setEnabled(False)
        self.btn_import.setEnabled(True)
        self.btn_mark_used.setEnabled(len(results) > 0)

        exact_count = sum(1 for r in results if r.is_exact)
        approx_count = len(results) - exact_count

        self.lbl_progress.setText(f"Complete! Found {len(results)} combinations ({exact_count} exact, {approx_count} approx)")
        self.status_bar.showMessage(f"Search complete. Found {len(results)} combinations.")

    def on_result_selected(self, item):
        """User clicked a result - highlight cells in Excel"""
        idx = item.data(Qt.ItemDataRole.UserRole)
        if idx is not None and idx < len(self.results):
            combination = self.results[idx]

            # Store currently selected indices
            self.current_selected_indices = combination.indices

            # Highlight in Excel
            self.excel_bridge.highlight_cells(combination.indices, self.highlight_color)

            self.status_bar.showMessage(f"Highlighted {len(combination.indices)} cells in Excel")
            self.btn_mark_used.setEnabled(True)

    def mark_cells_as_used(self):
        """Permanently color selected cells"""
        if not self.current_selected_indices:
            QMessageBox.information(self, "Info", "Please select a combination first")
            return

        # Color cells permanently
        self.excel_bridge.color_cells_permanent(self.current_selected_indices, self.permanent_color)

        # Mark item in list
        current_item = self.results_list.currentItem()
        if current_item:
            current_text = current_item.text()
            if not current_text.startswith("USED:"):
                current_item.setText("USED: " + current_text)
                current_item.setForeground(QColor(128, 128, 128))

        self.status_bar.showMessage(f"Marked {len(self.current_selected_indices)} cells as used")

    def clear_colors(self):
        """Clear all Excel cell colors"""
        self.excel_bridge.clear_highlighting()
        self.status_bar.showMessage("Cleared all cell colors in Excel")

    def filter_results(self, filter_type):
        """Filter displayed results"""
        for i in range(self.results_list.count()):
            item = self.results_list.item(i)
            text = item.text()

            if filter_type == 'all':
                item.setHidden(False)
            elif filter_type == 'exact':
                # Show only exact matches (dark green background)
                is_exact = item.background().color() == QColor(0, 100, 0)
                item.setHidden(not is_exact)
            elif filter_type == 'approx':
                # Show only approximate matches (dark orange background)
                is_approx = item.background().color() == QColor(139, 69, 19)
                item.setHidden(not is_approx)

    def pick_color(self, color_type):
        """Open color picker dialog"""
        color = QColorDialog.getColor()
        if color.isValid():
            hex_color = color.name()
            if color_type == 'highlight':
                self.highlight_color = hex_color
                self.btn_highlight_color.setStyleSheet(f"background-color: {hex_color}; border: 1px solid black;")
            else:
                self.permanent_color = hex_color
                self.btn_permanent_color.setStyleSheet(f"background-color: {hex_color}; border: 1px solid black;")


def main():
    app = QApplication(sys.argv)

    # Check if xlwings is available
    if not XLWINGS_AVAILABLE:
        QMessageBox.warning(None, "Missing Dependency",
                          "xlwings is not installed.\n\n"
                          "Please install it with:\npip install xlwings\n\n"
                          "The application will start but Excel features will be disabled.")

    window = MainWindow()
    window.show()

    sys.exit(app.exec())


if __name__ == '__main__':
    main()
