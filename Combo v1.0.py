import tkinter as tk
from tkinter import ttk, scrolledtext, messagebox
import threading
import queue
import time
import re
from collections import Counter

class CombinationSumGUI:
    def __init__(self, root):
        self.root = root
        self.root.title("Combination Sum Finder - Enhanced Version")
        self.root.geometry("1100x800")
        
        # Variables to store data
        self.numbers = []
        self.exact_combinations = []
        self.approx_combinations = []
        self.result_queue = queue.Queue()
        self.stop_event = threading.Event()
        self.is_running = False
        
        self.setup_ui()
        self.start_queue_check()
    
    def setup_ui(self):
        # Create canvas and scrollbar for scrolling support on small screens
        canvas = tk.Canvas(self.root, highlightthickness=0)
        scrollbar = ttk.Scrollbar(self.root, orient="vertical", command=canvas.yview)

        # Scrollable frame that will contain all content
        scrollable_frame = ttk.Frame(canvas, padding="10")

        # Configure canvas
        canvas.configure(yscrollcommand=scrollbar.set)

        # Pack scrollbar and canvas
        scrollbar.pack(side="right", fill="y")
        canvas.pack(side="left", fill="both", expand=True)

        # Create window in canvas for the scrollable frame
        canvas_frame = canvas.create_window((0, 0), window=scrollable_frame, anchor="nw")

        # Bind mouse wheel events for scrolling
        def on_mousewheel(event):
            canvas.yview_scroll(int(-1*(event.delta/120)), "units")

        def on_mousewheel_linux(event):
            canvas.yview_scroll(-1, "units")

        def on_mousewheel_linux_down(event):
            canvas.yview_scroll(1, "units")

        # Bind for Windows and Mac
        canvas.bind_all("<MouseWheel>", on_mousewheel)
        # Bind for Linux
        canvas.bind_all("<Button-4>", on_mousewheel_linux)
        canvas.bind_all("<Button-5>", on_mousewheel_linux_down)

        # Update scroll region when frame size changes
        def on_frame_configure(event):
            canvas.configure(scrollregion=canvas.bbox("all"))

        scrollable_frame.bind("<Configure>", on_frame_configure)

        # Make canvas expand to fill window width
        def on_canvas_configure(event):
            canvas.itemconfig(canvas_frame, width=event.width)

        canvas.bind("<Configure>", on_canvas_configure)

        # Main frame (now inside scrollable_frame)
        main_frame = scrollable_frame
        main_frame.columnconfigure(1, weight=1)
        
        # Input Section
        input_frame = ttk.LabelFrame(main_frame, text="Input Configuration", padding="10")
        input_frame.grid(row=0, column=0, columnspan=3, sticky=(tk.W, tk.E), pady=5)
        input_frame.columnconfigure(2, weight=1)
        
        # Target Sum and Buffer
        ttk.Label(input_frame, text="Target Sum:").grid(row=0, column=0, sticky=tk.W, pady=5)
        self.sum_entry = ttk.Entry(input_frame, width=15)
        self.sum_entry.grid(row=0, column=1, sticky=tk.W, pady=5, padx=(0,10))
        
        ttk.Label(input_frame, text="Buffer (±):").grid(row=0, column=2, sticky=tk.W, pady=5)
        self.buffer_entry = ttk.Entry(input_frame, width=10)
        self.buffer_entry.insert(0, "0")
        self.buffer_entry.grid(row=0, column=3, sticky=tk.W, pady=5, padx=(0,10))
        
        # Max Results and Max Length
        ttk.Label(input_frame, text="Max Results:").grid(row=0, column=4, sticky=tk.W, pady=5)
        self.max_results_entry = ttk.Entry(input_frame, width=10)
        self.max_results_entry.insert(0, "100")
        self.max_results_entry.grid(row=0, column=5, sticky=tk.W, pady=5)
        
        ttk.Label(input_frame, text="Max Length:").grid(row=0, column=6, sticky=tk.W, pady=5, padx=(10,0))
        self.max_length_entry = ttk.Entry(input_frame, width=10)
        self.max_length_entry.insert(0, "15")
        self.max_length_entry.grid(row=0, column=7, sticky=tk.W, pady=5)
        
        # Data format selection
        format_frame = ttk.Frame(input_frame)
        format_frame.grid(row=1, column=0, columnspan=6, sticky=tk.W, pady=10)
        
        ttk.Label(format_frame, text="Data Format:").pack(side=tk.LEFT)
        self.format_var = tk.StringVar(value="excel")
        ttk.Radiobutton(format_frame, text="📊 Copied from Excel (comma-separated)", 
                       variable=self.format_var, value="excel").pack(side=tk.LEFT, padx=10)
        ttk.Radiobutton(format_frame, text="📝 Numbers on separate lines", 
                       variable=self.format_var, value="lines").pack(side=tk.LEFT)
        
        # Numbers Input
        ttk.Label(input_frame, text="Numbers:").grid(row=1, column=0, sticky=(tk.W, tk.N), pady=5)
        self.numbers_text = scrolledtext.ScrolledText(input_frame, width=70, height=6)
        self.numbers_text.grid(row=1, column=1, columnspan=7, sticky=(tk.W, tk.E), pady=5)
        
        # Instructions
        instruction_text = ("Excel format: Copy cells and paste (preserves commas)\n"
                          "Line format: One number per line\n"
                          "Buffer example: Target=42, Buffer=2 finds sums from 40 to 44 (includes ±0, ±1, ±2)")
        instruction = ttk.Label(input_frame, text=instruction_text)
        instruction.grid(row=3, column=0, columnspan=6, sticky=tk.W, pady=5)
        
        # Control Buttons Frame
        control_frame = ttk.Frame(main_frame)
        control_frame.grid(row=1, column=0, columnspan=3, pady=10)
        
        self.find_button = ttk.Button(control_frame, text="🔍 Start Finding Combinations", command=self.start_finding)
        self.find_button.pack(side=tk.LEFT, padx=5)
        
        self.stop_button = ttk.Button(control_frame, text="⏹️ Stop Search", command=self.stop_finding, state="disabled")
        self.stop_button.pack(side=tk.LEFT, padx=5)
        
        self.clear_button = ttk.Button(control_frame, text="🗑️ Clear Results", command=self.clear_results)
        self.clear_button.pack(side=tk.LEFT, padx=5)
        
        # Debug test button
        self.test_button = ttk.Button(control_frame, text="🧪 Quick Test", command=self.run_quick_test)
        self.test_button.pack(side=tk.LEFT, padx=5)
        
        # Status Section
        status_frame = ttk.LabelFrame(main_frame, text="Search Status", padding="10")
        status_frame.grid(row=2, column=0, columnspan=3, sticky=(tk.W, tk.E), pady=5)
        
        self.status_label = ttk.Label(status_frame, text="Ready to search...")
        self.status_label.grid(row=0, column=0, sticky=tk.W)
        
        self.progress_label = ttk.Label(status_frame, text="")
        self.progress_label.grid(row=1, column=0, sticky=tk.W)
        
        # Results Summary
        summary_frame = ttk.Frame(status_frame)
        summary_frame.grid(row=2, column=0, sticky=tk.W, pady=5)
        
        self.exact_label = ttk.Label(summary_frame, text="Exact matches: 0", foreground="green")
        self.exact_label.pack(side=tk.LEFT, padx=(0,20))
        
        self.approx_label = ttk.Label(summary_frame, text="Approximate matches: 0", foreground="orange")
        self.approx_label.pack(side=tk.LEFT)
        
        # Results Section
        results_frame = ttk.LabelFrame(main_frame, text="Results", padding="10")
        results_frame.grid(row=3, column=0, columnspan=3, sticky=(tk.W, tk.E, tk.N, tk.S), pady=10)
        results_frame.columnconfigure(0, weight=1)
        results_frame.rowconfigure(3, weight=1)
        
        # Display controls
        display_frame = ttk.Frame(results_frame)
        display_frame.grid(row=0, column=0, sticky=(tk.W, tk.E), pady=5)
        
        ttk.Label(display_frame, text="Show:").pack(side=tk.LEFT)
        
        self.show_var = tk.StringVar(value="all")
        ttk.Radiobutton(display_frame, text="All", variable=self.show_var, value="all", command=self.update_display).pack(side=tk.LEFT, padx=5)
        ttk.Radiobutton(display_frame, text="First 10", variable=self.show_var, value="10", command=self.update_display).pack(side=tk.LEFT)
        ttk.Radiobutton(display_frame, text="First 25", variable=self.show_var, value="25", command=self.update_display).pack(side=tk.LEFT)
        ttk.Radiobutton(display_frame, text="Last 10", variable=self.show_var, value="last10", command=self.update_display).pack(side=tk.LEFT)
        
        ttk.Label(display_frame, text="Custom:").pack(side=tk.LEFT, padx=(10,0))
        self.custom_display_entry = ttk.Entry(display_frame, width=5)
        self.custom_display_entry.pack(side=tk.LEFT, padx=5)
        ttk.Button(display_frame, text="Update", command=self.update_display).pack(side=tk.LEFT)
        
        # Sorting controls
        sort_frame = ttk.Frame(results_frame)
        sort_frame.grid(row=1, column=0, sticky=(tk.W, tk.E), pady=5)
        
        ttk.Label(sort_frame, text="🔢 Sort by:").pack(side=tk.LEFT)
        self.sort_var = tk.StringVar(value="length")
        ttk.Radiobutton(sort_frame, text="📏 Length (Shortest First)", variable=self.sort_var, value="length", command=self.update_display).pack(side=tk.LEFT, padx=5)
        ttk.Radiobutton(sort_frame, text="🔍 Found Order", variable=self.sort_var, value="found", command=self.update_display).pack(side=tk.LEFT)
        ttk.Radiobutton(sort_frame, text="📊 Sum Value", variable=self.sort_var, value="sum", command=self.update_display).pack(side=tk.LEFT)
        
        # Original numbers display
        original_frame = ttk.LabelFrame(results_frame, text="Original Numbers (click combinations below to highlight)", padding="5")
        original_frame.grid(row=2, column=0, sticky=(tk.W, tk.E), pady=5)
        original_frame.columnconfigure(0, weight=1)
        
        self.original_text = tk.Text(original_frame, height=3, wrap=tk.WORD)
        self.original_text.grid(row=0, column=0, sticky=(tk.W, tk.E))
        
        # Results Notebook for exact vs approximate
        self.results_notebook = ttk.Notebook(results_frame)
        self.results_notebook.grid(row=3, column=0, sticky=(tk.W, tk.E, tk.N, tk.S), pady=5)
        
        # Exact matches tab
        exact_frame = ttk.Frame(self.results_notebook)
        self.results_notebook.add(exact_frame, text="🎯 Exact Matches")
        exact_frame.columnconfigure(0, weight=1)
        exact_frame.rowconfigure(0, weight=1)
        
        self.exact_text = scrolledtext.ScrolledText(exact_frame, height=10)
        self.exact_text.grid(row=0, column=0, sticky=(tk.W, tk.E, tk.N, tk.S))
        self.exact_text.bind("<Button-1>", lambda e: self.on_combination_click(e, "exact"))
        
        # Approximate matches tab
        approx_frame = ttk.Frame(self.results_notebook)
        self.results_notebook.add(approx_frame, text="🎭 Approximate Matches")
        approx_frame.columnconfigure(0, weight=1)
        approx_frame.rowconfigure(0, weight=1)
        
        self.approx_text = scrolledtext.ScrolledText(approx_frame, height=10)
        self.approx_text.grid(row=0, column=0, sticky=(tk.W, tk.E, tk.N, tk.S))
        self.approx_text.bind("<Button-1>", lambda e: self.on_combination_click(e, "approx"))
        
        # Configure text highlighting
        self.original_text.tag_configure("highlight", background="yellow", foreground="black")
        self.exact_text.tag_configure("selected_line", background="#e6f3ff")
        self.approx_text.tag_configure("selected_line", background="#ffe6e6")
        
        # Configure grid weights for resizing
        main_frame.rowconfigure(3, weight=1)
    
    def parse_numbers(self, text):
        """Parse numbers based on selected format"""
        if self.format_var.get() == "excel":
            # Excel format: comma-separated, may have tabs/spaces
            text = re.sub(r'[\t\n\r]+', ' ', text.strip())
            text = re.sub(r'\s+', ' ', text)  # Normalize spaces
            
            numbers = []
            for item in text.replace(',', ' ').split():
                try:
                    num = float(item.strip())
                    if num.is_integer():
                        numbers.append(int(num))
                    else:
                        numbers.append(num)
                except ValueError:
                    continue
        else:
            # Line format: one number per line
            lines = text.strip().split('\n')
            numbers = []
            for line in lines:
                line = line.strip()
                if line:
                    try:
                        num = float(line)
                        if num.is_integer():
                            numbers.append(int(num))
                        else:
                            numbers.append(num)
                    except ValueError:
                        continue
        
        return numbers
    
    def find_combinations_simple(self, numbers, target, buffer, max_results, max_length, start_time):
        """Super simple approach that finds SHORTEST combinations first"""
        # Sort in DESCENDING order (largest numbers first) for shorter combinations
        numbers.sort(reverse=True)
        results = []
        count = 0

        def find_recursive(start_idx, current_sum, current_combo):
            nonlocal count
            if count >= max_results:
                return
            if self.stop_event.is_set():
                return
            
            # CRITICAL: Stop if combination gets too long
            if len(current_combo) > max_length:
                return
                
            # Check if current sum is within buffer range
            if current_combo:  # Don't check empty combination
                diff = abs(current_sum - target)
                if diff <= buffer:
                    if current_sum == target:
                        results.append(('exact', current_combo.copy(), current_sum))
                    else:
                        results.append(('approx', current_combo.copy(), current_sum))
                    count += 1

                    # Send result immediately for streaming
                    elapsed = time.time() - start_time  # Real elapsed time
                    exact_count = sum(1 for r in results if r[0] == 'exact')
                    approx_count = count - exact_count
                    
                    self.result_queue.put({
                        'type': 'combination',
                        'result_type': current_sum == target and 'exact' or 'approx',
                        'data': current_combo.copy(),
                        'actual_sum': current_sum,
                        'exact_count': exact_count,
                        'approx_count': approx_count,
                        'elapsed': elapsed
                    })
                    
                    if count >= max_results:
                        return
            
            # Try adding each remaining number
            for i in range(start_idx, len(numbers)):
                if count >= max_results or self.stop_event.is_set():
                    return
                
                # STOP if combination would be too long
                if len(current_combo) >= max_length:
                    return
                    
                new_sum = current_sum + numbers[i]
                
                # Skip if way too big
                if new_sum > target + buffer:
                    continue  # Don't break since we're going descending

                # Skip duplicates at same level to avoid duplicate combinations
                # This is correct: it only skips when we're at the same recursion level (i > start_idx)
                # which prevents exploring the same subproblem twice with different instances of the same number
                # Example: [1a, 1b, 2] - at level 0, we try 1a, then skip 1b (would create same subtree)
                # But when recursing from 1a, we can still use 1b at the next level
                if i > start_idx and abs(numbers[i] - numbers[i - 1]) < 1e-9:
                    continue
                
                # Recursive call
                current_combo.append(numbers[i])
                find_recursive(i + 1, new_sum, current_combo)
                current_combo.pop()
        
        # Start the search
        find_recursive(0, 0, [])
        return results
    
    def worker_thread(self, numbers, target, buffer, max_results, max_length):
        """Worker thread that finds combinations and puts them in queue"""
        try:
            start_time = time.time()

            # Use the simple recursive method with immediate streaming
            self.find_combinations_simple(numbers, target, buffer, max_results, max_length, start_time)
            
            # Send completion signal
            elapsed = time.time() - start_time
            exact_count = len(self.exact_combinations)
            approx_count = len(self.approx_combinations)
            
            self.result_queue.put({
                'type': 'complete',
                'exact_count': exact_count,
                'approx_count': approx_count,
                'elapsed': elapsed
            })
            
        except Exception as e:
            self.result_queue.put({
                'type': 'error',
                'error': str(e)
            })
    
    def start_queue_check(self):
        """Start the queue checking process"""
        self.check_queue()
    
    def check_queue(self):
        """Check for results from worker thread"""
        try:
            while True:
                result = self.result_queue.get_nowait()
                
                if result['type'] == 'combination':
                    # New combination found
                    if result['result_type'] == 'exact':
                        self.exact_combinations.append((result['data'], result['actual_sum']))
                    else:
                        self.approx_combinations.append((result['data'], result['actual_sum']))
                    
                    # Update labels
                    self.exact_label.config(text=f"Exact matches: {result['exact_count']}")
                    self.approx_label.config(text=f"Approximate matches: {result['approx_count']}")
                    total_count = result['exact_count'] + result['approx_count']
                    
                    # Fix division by zero error
                    if result['elapsed'] > 0:
                        rate = total_count / result['elapsed']
                        self.progress_label.config(text=f"Time: {result['elapsed']:.2f}s | Rate: {rate:.1f}/sec")
                    else:
                        self.progress_label.config(text=f"Time: {result['elapsed']:.2f}s | Rate: Very Fast!")
                    
                    # Update display if showing all results and not sorting by length
                    if self.show_var.get() == "all" and self.sort_var.get() == "found":
                        self.append_combination_to_display(result['result_type'], result['data'], result['actual_sum'])
                    else:
                        # Update display periodically for other modes or when sorting
                        if total_count % 10 == 0:
                            self.update_display()
                
                elif result['type'] == 'complete':
                    # Search completed
                    self.search_completed(result['exact_count'], result['approx_count'], result['elapsed'])
                
                elif result['type'] == 'error':
                    # Error occurred
                    messagebox.showerror("Error", f"An error occurred: {result['error']}")
                    self.search_completed(0, 0, 0)
                
        except queue.Empty:
            pass
        
        # Schedule next check
        self.root.after(50, self.check_queue)
    
    def append_combination_to_display(self, result_type, combination, actual_sum):
        """Append a single combination to the appropriate display"""
        combo_str = "{" + ", ".join(map(str, combination)) + "}"
        
        if result_type == 'exact':
            count = len(self.exact_combinations)
            text_widget = self.exact_text
            display_str = f"{count}. {combo_str}\n"
        else:
            count = len(self.approx_combinations)
            text_widget = self.approx_text
            diff = actual_sum - float(self.sum_entry.get())
            sign = "+" if diff > 0 else ""
            display_str = f"{count}. {combo_str} = {actual_sum} ({sign}{diff:.2f})\n"
        
        text_widget.insert(tk.END, display_str)
        text_widget.see(tk.END)  # Auto-scroll to bottom
    
    def start_finding(self):
        """Start the combination finding process"""
        try:
            # Validate inputs
            target = float(self.sum_entry.get())
            buffer = float(self.buffer_entry.get())
            max_results = int(self.max_results_entry.get())
            max_length = int(self.max_length_entry.get())
            
            if max_results <= 0:
                messagebox.showerror("Error", "Max results must be a positive number")
                return
            
            if max_length <= 0:
                messagebox.showerror("Error", "Max length must be a positive number")
                return
            
            if buffer < 0:
                messagebox.showerror("Error", "Buffer must be non-negative")
                return
            
            # Get and parse numbers
            numbers_text = self.numbers_text.get("1.0", tk.END).strip()
            if not numbers_text:
                messagebox.showerror("Error", "Please enter some numbers")
                return
            
            self.numbers = self.parse_numbers(numbers_text)
            if not self.numbers:
                messagebox.showerror("Error", "No valid numbers found")
                return
            
            # Stop any existing search and reset state
            if self.is_running:
                self.stop_event.set()
                # Thread will stop on next iteration - no need to block UI
            
            # Clear previous results
            self.exact_combinations = []
            self.approx_combinations = []
            self.exact_text.delete("1.0", tk.END)
            self.approx_text.delete("1.0", tk.END)
            self.exact_label.config(text="Exact matches: 0")
            self.approx_label.config(text="Approximate matches: 0")
            self.progress_label.config(text="")
            
            # Display original numbers
            self.display_original_numbers()
            
            # Update UI state
            self.is_running = True
            self.find_button.config(state="disabled")
            self.stop_button.config(state="normal")
            
            if buffer > 0:
                self.status_label.config(text=f"🔍 Searching: target={target} ±{buffer} | Numbers: {len(self.numbers)} | Max Length: {max_length}")
            else:
                self.status_label.config(text=f"🔍 Searching for exact matches: target={target} | Numbers: {len(self.numbers)} | Max Length: {max_length}")
            
            # Reset stop event and start worker thread
            self.stop_event.clear()
            worker = threading.Thread(target=self.worker_thread, args=(self.numbers, target, buffer, max_results, max_length))
            worker.daemon = True
            worker.start()
            
        except ValueError as e:
            messagebox.showerror("Error", "Please enter valid numbers for all fields")
        except Exception as e:
            messagebox.showerror("Error", f"An error occurred: {str(e)}")
    
    def stop_finding(self):
        """Stop the combination finding process"""
        if self.is_running:
            self.stop_event.set()
            self.status_label.config(text="🛑 Stopping search...")
            # Don't change button states here - let search_completed handle it
    
    def search_completed(self, exact_count, approx_count, elapsed):
        """Handle search completion"""
        # Check if search was stopped BEFORE clearing the event
        was_stopped = self.stop_event.is_set()

        # Reset threading state
        self.is_running = False
        self.stop_event.clear()

        # Reset button states
        self.find_button.config(state="normal")
        self.stop_button.config(state="disabled")

        total = exact_count + approx_count
        if was_stopped:
            self.status_label.config(text=f"🛑 Search stopped. Found {total} combinations ({exact_count} exact, {approx_count} approx) in {elapsed:.2f}s")
        else:
            self.status_label.config(text=f"✅ Search completed! Found {total} combinations ({exact_count} exact, {approx_count} approx) in {elapsed:.2f}s")
        
        # Final display update
        self.update_display()
    
    def clear_results(self):
        """Clear all results and reset application state"""
        # Stop any running search first
        if self.is_running:
            self.stop_event.set()
            # Thread will stop on next iteration - no need to block UI
        
        # Reset threading state
        self.is_running = False
        self.stop_event.clear()
        
        # Reset button states
        self.find_button.config(state="normal")
        self.stop_button.config(state="disabled")
        
        # Clear all data
        self.exact_combinations = []
        self.approx_combinations = []
        
        # Clear all text widgets
        self.exact_text.delete("1.0", tk.END)
        self.approx_text.delete("1.0", tk.END)
        self.original_text.delete("1.0", tk.END)
        
        # Reset labels
        self.exact_label.config(text="Exact matches: 0")
        self.approx_label.config(text="Approximate matches: 0")
        self.progress_label.config(text="")
        self.status_label.config(text="Ready to search...")
        
        # Clear any remaining items in the queue
        try:
            while True:
                self.result_queue.get_nowait()
        except queue.Empty:
            pass
    
    def display_original_numbers(self):
        """Display original numbers in the text widget"""
        self.original_text.delete("1.0", tk.END)
        if self.format_var.get() == "excel":
            numbers_str = ", ".join(map(str, self.numbers))
        else:
            numbers_str = "\n".join(map(str, self.numbers))
        self.original_text.insert("1.0", numbers_str)
    
    def sort_combinations(self, combinations):
        """Sort combinations based on selected criteria"""
        if not combinations:
            return combinations
            
        sort_option = self.sort_var.get()
        
        if sort_option == "length":
            # Sort by length (shortest first), then by sum for ties
            return sorted(combinations, key=lambda x: (len(x[0]), x[1]))
        elif sort_option == "sum":
            # Sort by sum value
            return sorted(combinations, key=lambda x: x[1])
        else:  # "found" - original order
            return combinations
    
    def update_display(self):
        """Update the combinations display based on selected option"""
        # Clear current displays
        self.exact_text.delete("1.0", tk.END)
        self.approx_text.delete("1.0", tk.END)
        
        # Sort combinations first
        sorted_exact = self.sort_combinations(self.exact_combinations)
        sorted_approx = self.sort_combinations(self.approx_combinations)
        
        # Get display parameters
        show_option = self.show_var.get()
        
        # Update exact matches with grouping by length if sorted by length
        exact_to_show = self.get_combinations_to_show(sorted_exact, show_option)
        self.display_combinations_with_grouping(self.exact_text, exact_to_show, "exact")
        
        # Update approximate matches
        approx_to_show = self.get_combinations_to_show(sorted_approx, show_option)
        self.display_combinations_with_grouping(self.approx_text, approx_to_show, "approx")
    
    def display_combinations_with_grouping(self, text_widget, combinations, combo_type):
        """Display combinations with optional grouping by length"""
        if not combinations:
            return
            
        current_length = None
        combo_number = 1
        target = float(self.sum_entry.get())
        
        for combination, actual_sum in combinations:
            combo_length = len(combination)
            
            # Add length header if sorting by length and length changes
            if self.sort_var.get() == "length" and combo_length != current_length:
                if current_length is not None:  # Not the first group
                    text_widget.insert(tk.END, "\n")
                
                length_text = "number" if combo_length == 1 else "numbers"
                text_widget.insert(tk.END, f"═══ {combo_length} {length_text} ═══\n")
                current_length = combo_length
            
            # Format the combination display
            combo_str = "{" + ", ".join(map(str, combination)) + "}"
            
            if combo_type == "exact":
                display_str = f"{combo_number}. {combo_str}\n"
            else:
                diff = actual_sum - target
                sign = "+" if diff > 0 else ""
                display_str = f"{combo_number}. {combo_str} = {actual_sum} ({sign}{diff:.2f})\n"
            
            text_widget.insert(tk.END, display_str)
            combo_number += 1
    
    def get_combinations_to_show(self, combinations, show_option):
        """Get the combinations to display based on show option"""
        if not combinations:
            return []
        
        if show_option == "all":
            return combinations
        elif show_option == "last10":
            return combinations[-10:]
        else:
            try:
                num_to_show = int(show_option)
            except ValueError:
                try:
                    num_to_show = int(self.custom_display_entry.get())
                except ValueError:
                    return combinations
            
            return combinations[:num_to_show]
    
    def on_combination_click(self, event, tab_type):
        """Handle click on combination to highlight numbers"""
        text_widget = self.exact_text if tab_type == "exact" else self.approx_text
        
        # Get the line that was clicked
        index = text_widget.index(tk.CURRENT)
        line_start = index.split('.')[0] + '.0'
        line_end = index.split('.')[0] + '.end'
        
        # Clear previous selection
        text_widget.tag_remove("selected_line", "1.0", tk.END)
        
        # Highlight selected line
        text_widget.tag_add("selected_line", line_start, line_end)
        
        # Extract combination from the line
        line_content = text_widget.get(line_start, line_end)
        
        # Parse the combination numbers
        try:
            # Extract the part between { and }
            start = line_content.find('{')
            end = line_content.find('}')
            if start != -1 and end != -1:
                combo_str = line_content[start+1:end]
                combination = []
                for x in combo_str.split(','):
                    x = x.strip()
                    if x:
                        try:
                            if '.' in x:
                                combination.append(float(x))
                            else:
                                combination.append(int(x))
                        except ValueError:
                            continue
                
                self.highlight_numbers_with_duplicates(combination)
        except:
            pass  # If parsing fails, just ignore
    
    def highlight_numbers_with_duplicates(self, combination):
        """Highlight numbers in original list, handling duplicates correctly"""
        # Clear previous highlights
        self.original_text.tag_remove("highlight", "1.0", tk.END)
        
        # Count how many times each number appears in the combination
        combination_counts = Counter(combination)
        
        # Get the original text content
        original_content = self.original_text.get("1.0", tk.END)
        
        # For each unique number in combination, find and highlight the required instances
        for number, required_count in combination_counts.items():
            self.highlight_number_instances(number, required_count)
    
    def highlight_number_instances(self, number, count_needed):
        """Highlight exactly 'count_needed' instances of 'number' in the original text"""
        num_str = str(number)
        found_count = 0
        start_pos = "1.0"
        
        while found_count < count_needed:
            # Search for the next occurrence
            pos = self.original_text.search(num_str, start_pos, tk.END)
            if not pos:
                break  # No more instances found
            
            end_pos = f"{pos}+{len(num_str)}c"
            
            # Check if this is a word boundary match (complete number, not part of another number)
            char_before = self.original_text.get(f"{pos}-1c", pos) if pos != "1.0" else ""
            char_after = self.original_text.get(end_pos, f"{end_pos}+1c")
            
            # Valid separators: start/end of text, space, comma, tab, newline
            valid_separators = ["", " ", ",", "\t", "\n"]
            
            if char_before in valid_separators and char_after in valid_separators:
                # This is a complete number match - highlight it
                self.original_text.tag_add("highlight", pos, end_pos)
                found_count += 1

            # Move to the character after the current match to continue searching
            start_pos = f"{pos}+1c"
    
    def run_quick_test(self):
        """Run a quick test to verify the algorithm works"""
        # Set simple test values
        self.sum_entry.delete(0, tk.END)
        self.sum_entry.insert(0, "350")
        
        self.buffer_entry.delete(0, tk.END)
        self.buffer_entry.insert(0, "0")
        
        self.max_results_entry.delete(0, tk.END)
        self.max_results_entry.insert(0, "20")
        
        self.max_length_entry.delete(0, tk.END)
        self.max_length_entry.insert(0, "10")
        
        self.numbers_text.delete("1.0", tk.END)
        self.numbers_text.insert("1.0", "7, 14, 21, 28, 35, 42, 49, 7, 14, 21, 28, 35, 42, 49")
        
        self.format_var.set("excel")
        self.sort_var.set("length")  # Default to length sorting
        
        # Show what we're testing
        messagebox.showinfo("Quick Test", "Testing with YOUR example:\nTarget: 350 (exact)\nNumbers: 7,14,21,28,35,42,49 (repeated)\nMax Length: 10\n\nShould find:\n• {49,49,49,49,49,49,49,7} (8 numbers) first!\n• NOT {7×50} (50 numbers)\n\nSorted LARGEST FIRST!")
        
        # Start the test
        self.start_finding()

def main():
    root = tk.Tk()
    app = CombinationSumGUI(root)
    root.mainloop()

if __name__ == "__main__":
    main()