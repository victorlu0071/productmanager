class PasteboardMonitor:
    def __init__(self, entry_field, label, status_label):
        """
        Initializes the pasteboard monitor.
        
        Args:
            entry_field (tk.Entry): The input entry field where clipboard content is pasted.
            label (tk.Label): The label indicating the current input prompt (e.g., 'Enter Product Name:').
            status_label (tk.Label): The status label for displaying errors or status messages.
        """
        self.entry_field = entry_field
        self.label = label
        self.status_label = status_label
        self.last_clipboard_content = None
        self.monitoring = False

    def start_monitoring(self):
        """Starts the clipboard monitoring in a separate thread."""
        if not self.monitoring:
            self.monitoring = True
            self.monitor_thread = threading.Thread(target=self.monitor_clipboard, daemon=True)
            self.monitor_thread.start()

    def stop_monitoring(self):
        """Stops the clipboard monitoring."""
        self.monitoring = False

    def monitor_clipboard(self):
        """Continuously monitors the clipboard for changes."""
        while self.monitoring:
            try:
                current_clipboard_content = self.get_clipboard_content()
                if current_clipboard_content and current_clipboard_content != self.last_clipboard_content:
                    self.last_clipboard_content = current_clipboard_content
                    self.handle_clipboard_content(current_clipboard_content)
            except Exception as e:
                self.status_label.config(text=f"Error monitoring clipboard: {e}")
            time.sleep(0.5)  # Check for clipboard changes every 0.5 seconds

    def get_clipboard_content(self):
        """Fetches the current content from the system clipboard."""
        try:
            root = tk.Tk()
            root.withdraw()
            clipboard_content = root.clipboard_get()
            root.destroy()
            return clipboard_content.strip()
        except Exception:
            return None

    def handle_clipboard_content(self, content):
        """
        Handles the clipboard content by pasting it into the input field if the prompt 
        asks for 'Product Name' or 'Product Link', then triggers an 'Enter' event.
        
        Args:
            content (str): The content from the clipboard.
        """
        current_prompt = self.label.cget("text").strip().lower()
        if "product name" in current_prompt or "product link" in current_prompt:
            self.entry_field.delete(0, tk.END)
            self.entry_field.insert(0, content)
            self.simulate_enter_keypress()

    def simulate_enter_keypress(self):
        """Simulates the 'Enter' key press to process the input automatically."""
        self.entry_field.event_generate('<Return>')