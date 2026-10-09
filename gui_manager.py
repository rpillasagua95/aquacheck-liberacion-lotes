# -*- coding: utf-8 -*-
# gui_manager.py (MODIFICADO CON MEJORAS DE UI/UX)

import tkinter as tk
from tkinter import filedialog, scrolledtext, messagebox, ttk
import ttkbootstrap as ttkb
from ttkbootstrap.constants import *
from PIL import Image, ImageTk
import os
import threading
import webbrowser
from tkinterdnd2 import DND_FILES
import openpyxl
import io
from contextlib import redirect_stdout

# Importaciones de módulos locales
import backend_logic as be
import config_manager as cm
import whatsapp_service as ws

class AppGUI:
    def __init__(self, root):
        self.root = root
        self.root.title("AquaCheck · Liberación de lotes en recepción")
        self.root.geometry("1320x880")
        self.root.minsize(900, 700)
        self.sabor_alert_label = None
        self.selected_additional_photo_path = tk.StringVar()
        self.screenshot_path = None
        
        self.residual_check_var = tk.BooleanVar()
        self.current_analysis_data = {}
        self.sample_status_widgets = {}
        self.selected_excel_file = None
        self.image_task_buttons = []
        
        self.whatsapp_is_authorized = False
        self.select_folder_button = None
        
        # Nuevas propiedades para mejoras
        self.current_lote = None
        self.sample_buttons_container = None
        self.alerts_tree = None
        self.delete_photo_button = None
        self.sample_canvas = None

        self.create_main_layout()
        self.create_notebook()
        self.create_statusbar()
        
        self.update_whatsapp_status()
        self.check_initial_config()

    def apply_modern_style(self):
        """Tipografia y estilos comunes: Segoe UI, titulos de seccion en negrita, pestañas amplias."""
        style = ttkb.Style()
        base = ("Segoe UI", 10)
        self.root.option_add("*Font", base)
        style.configure(".", font=base)
        style.configure("TLabelframe.Label", font=("Segoe UI Semibold", 10))
        style.configure("TNotebook.Tab", padding=(16, 8), font=("Segoe UI", 10))
        style.configure("TButton", padding=(10, 6))
        colors = style.colors
        self.palette = {"bg": colors.bg, "primary": colors.primary, "fg": colors.fg,
                        "muted": colors.secondary, "select": colors.info}

    def create_header(self):
        """Cabecera con el nombre de la app y lo que hace."""
        header = ttkb.Frame(self.root, bootstyle="primary", padding=(18, 12))
        header.pack(side=TOP, fill=X)
        ttkb.Label(header, text="AquaCheck", font=("Segoe UI Semibold", 18), bootstyle="inverse-primary").pack(side=LEFT)
        ttkb.Label(header, text="   Liberación de lotes en recepción · reporte por WhatsApp",
                   font=("Segoe UI", 10), bootstyle="inverse-primary").pack(side=LEFT, pady=(6, 0))

    def create_main_layout(self):
        self.apply_modern_style()
        self.create_header()
        main_frame = ttkb.Frame(self.root, padding=(14, 12, 14, 6))
        main_frame.pack(fill=BOTH, expand=True)
        self.main_frame = main_frame
        self.notebook = ttkb.Notebook(main_frame)
        self.notebook.pack(fill=BOTH, expand=True)

    def create_notebook(self):
        self.liberacion_tab = ttkb.Frame(self.notebook, padding="10")
        self.create_liberacion_widgets()
        self.notebook.add(self.liberacion_tab, text="  Liberación en Recepción  ", padding=5)
        
        self.imagenes_tab = ttkb.Frame(self.notebook, padding="10")
        self.create_imagenes_widgets()
        self.notebook.add(self.imagenes_tab, text="  Gestión de Imágenes en Excel  ", padding=5)
        
        self.config_tab = ttkb.Frame(self.notebook, padding="10")
        self.create_config_widgets()
        self.notebook.add(self.config_tab, text="  Configuración  ", padding=5)

    def create_statusbar(self):
        self.status_bar = ttkb.Frame(self.root, padding=(12, 6), bootstyle="light")
        # Se coloca antes del contenido para que nunca quede fuera de la ventana
        self.status_bar.pack(side=BOTTOM, fill=X, before=self.main_frame)
        self.status_label = ttkb.Label(self.status_bar, text="Iniciando...")
        self.status_label.pack(side=LEFT, padx=5)
        self.whatsapp_status_label = ttkb.Label(self.status_bar, text="WhatsApp: Verificando...")
        self.whatsapp_status_label.pack(side=RIGHT, padx=5)

    def create_liberacion_widgets(self):
        frame = self.liberacion_tab
        frame.columnconfigure(0, weight=3, minsize=330)
        frame.columnconfigure(1, weight=3, minsize=300)
        frame.columnconfigure(2, weight=5, minsize=400)
        frame.rowconfigure(0, weight=1)
        
        # --- Frame de Controles Izquierdo ---
        controls_frame = ttkb.Labelframe(frame, text="1. Selección de Archivo", padding="10")
        controls_frame.grid(row=0, column=0, sticky="nsew", padx=(0, 10))
        
        self.select_folder_button = ttkb.Button(controls_frame, text="📁 Seleccionar Carpeta", command=self.select_excel_folder, bootstyle=PRIMARY)
        self.select_folder_button.pack(fill=X, pady=(0,5))
        
        list_frame = ttkb.Frame(controls_frame)
        list_frame.pack(fill=BOTH, expand=True, pady=5)
        self.excel_files_list = tk.Listbox(list_frame, height=10, exportselection=False, font=("Segoe UI", 10),
                                           relief="flat", borderwidth=0, highlightthickness=1, activestyle="none",
                                           highlightcolor=self.palette["primary"], highlightbackground="#dee2e6",
                                           selectbackground=self.palette["primary"], selectforeground="#ffffff")
        self.excel_files_list.pack(side=LEFT, fill=BOTH, expand=True)

        scrollbar = ttkb.Scrollbar(list_frame, orient=VERTICAL, command=self.excel_files_list.yview, bootstyle="round")
        scrollbar.pack(side=RIGHT, fill=Y)
        self.excel_files_list.config(yscrollcommand=scrollbar.set)
        
        self.excel_files_list.bind("<<ListboxSelect>>", self.on_excel_file_select)
        self.excel_files_list.drop_target_register(DND_FILES)
        self.excel_files_list.dnd_bind('<<Drop>>', self.handle_file_drop)
        
        options_frame = ttkb.Labelframe(controls_frame, text="Opciones de Envío", padding=10)
        options_frame.pack(fill=BOTH, pady=10, expand=False)
        
        self.residual_check = ttkb.Checkbutton(options_frame, text="Incluir 'RESIDUAL EN ANÁLISIS'", variable=self.residual_check_var, bootstyle=INFO, command=self.on_residual_check_toggle)
        self.residual_check.pack(fill=X, pady=(5,0))
        
        photo_buttons_frame = ttkb.Frame(options_frame)
        photo_buttons_frame.pack(fill=X, pady=5)
        ttkb.Button(photo_buttons_frame, text="📎 Adjuntar Foto Adicional", command=self.select_additional_photo, bootstyle=SECONDARY).pack(side=LEFT, expand=True, fill=X)
        self.delete_photo_button = ttkb.Button(photo_buttons_frame, text="❌", command=self.clear_additional_photo, bootstyle="danger-outline", state="disabled")
        self.delete_photo_button.pack(side=LEFT, padx=(5,0))

        self.additional_photo_label = ttkb.Label(options_frame, text="Ninguna foto seleccionada.", wraplength=280, bootstyle="secondary")
        self.additional_photo_label.pack(fill=X)
        self.additional_photo_thumbnail = ttkb.Label(options_frame, text="⬇  Arrastre una imagen aquí", anchor="center", padding=18, borderwidth=1, relief="solid", bootstyle="secondary")
        self.additional_photo_thumbnail.pack(fill=X, pady=5)
        self.additional_photo_thumbnail.bind("<Button-1>", lambda e: self.show_image_in_new_window(self.selected_additional_photo_path.get()))
        
        self.additional_photo_thumbnail.drop_target_register(DND_FILES)
        self.additional_photo_thumbnail.dnd_bind('<<Drop>>', self.handle_image_drop)
        self.sabor_alert_label = ttkb.Label(options_frame, text="LOTE PRESENTA SABOR", bootstyle=DANGER, font=("-size 11 -weight bold"), anchor="center")
        
        # --- Frame Central ---
        center_frame = ttkb.Frame(frame)
        center_frame.grid(row=0, column=1, sticky="nsew", padx=(0, 10))
        # MEJORA: Panel de muestras con tamaño automático, el mensaje se expande
        center_frame.rowconfigure(0, weight=0) # Fila de muestras no se expande verticalmente
        center_frame.rowconfigure(1, weight=1) # Fila de mensaje sí se expande para llenar espacio
        center_frame.columnconfigure(0, weight=1)

        sample_labelframe = ttkb.Labelframe(center_frame, text="2. Seleccione una muestra para analizar", padding=10)
        sample_labelframe.grid(row=0, column=0, sticky="new", pady=(0, 10))
        
        # MEJORA: Añadido Canvas con scroll horizontal para muestras. Mantiene la altura compacta.
        self.sample_canvas = tk.Canvas(sample_labelframe, borderwidth=0, highlightthickness=0)
        sample_scrollbar = ttkb.Scrollbar(sample_labelframe, orient="horizontal", command=self.sample_canvas.xview, bootstyle="round")
        self.sample_buttons_container = ttkb.Frame(self.sample_canvas)

        self.sample_buttons_container.bind(
            "<Configure>", 
            lambda e: self.sample_canvas.configure(scrollregion=self.sample_canvas.bbox("all"))
        )
        self.sample_canvas.create_window((0, 0), window=self.sample_buttons_container, anchor="nw")
        self.sample_canvas.configure(xscrollcommand=sample_scrollbar.set)
        
        self.sample_canvas.pack(side="top", fill="x", expand=True)
        sample_scrollbar.pack(side="bottom", fill="x")

        message_frame = ttkb.Labelframe(center_frame, text="3. Verifique el Mensaje (Editable)", padding="10")
        message_frame.grid(row=1, column=0, sticky="nsew")
        self.message_text = scrolledtext.ScrolledText(message_frame, wrap=tk.WORD, height=10, font=("Segoe UI", 10))
        self.message_text.pack(fill=BOTH, expand=True)
        
        # --- Frame Derecho ---
        right_frame = ttkb.Frame(frame)
        right_frame.grid(row=0, column=2, sticky="nsew")
        right_frame.rowconfigure(0, weight=3)
        right_frame.rowconfigure(1, weight=2)
        right_frame.columnconfigure(0, weight=1)
        
        preview_frame = ttkb.Labelframe(right_frame, text="4. Vista Previa (Clic para agrandar)", padding="10")
        preview_frame.grid(row=0, column=0, sticky="nsew", pady=(0, 5))
        preview_frame.rowconfigure(0, weight=1)
        preview_frame.columnconfigure(0, weight=1)
        
        self.image_preview_label = ttkb.Label(preview_frame, text="La captura del Excel aparecerá aquí.", anchor="center")
        self.image_preview_label.grid(row=0, column=0, sticky="nsew")
        self.image_preview_label.bind("<Button-1>", lambda e: self.show_image_in_new_window(self.screenshot_path))
        self.image_preview_label.config(cursor="hand2")
        
        self.send_controls_frame = ttkb.Frame(preview_frame)
        self.send_controls_frame.grid(row=1, column=0, sticky="ew", pady=(10, 0))
        self.send_controls_frame.columnconfigure(0, weight=1)
        
        self.send_button = ttkb.Button(self.send_controls_frame, text="✔️ Enviar a WhatsApp", command=self.send_to_whatsapp, bootstyle=SUCCESS, state="disabled")
        self.send_button.grid(row=0, column=0, sticky="ew", ipady=5)
        
        self.send_progressbar = ttkb.Progressbar(self.send_controls_frame, mode='indeterminate', bootstyle=SUCCESS)
        self.whatsapp_indicator = ttkb.Label(self.send_controls_frame, text="⚪", font=("Segoe UI", 14, "bold"), bootstyle=SECONDARY, width=2, anchor="center")
        self.whatsapp_indicator.grid(row=0, column=1, padx=(10, 0))
        
        alerts_frame = ttkb.Labelframe(right_frame, text="Advertencias (Doble clic para expandir)", padding="10", bootstyle=WARNING)
        alerts_frame.grid(row=1, column=0, sticky="nsew")
        
        cols = ("tipo", "prioridad", "descripcion")
        self.alerts_tree = ttk.Treeview(alerts_frame, columns=cols, show='headings', height=5, selectmode="none")
        
        self.alerts_tree.heading("tipo", text="Tipo")
        self.alerts_tree.heading("prioridad", text="Prioridad")
        self.alerts_tree.heading("descripcion", text="Descripción")
        
        self.alerts_tree.column("tipo", width=70, stretch=False, anchor="center")
        self.alerts_tree.column("prioridad", width=90, stretch=False, anchor="center")
        self.alerts_tree.column("descripcion", width=300)

        tree_scrollbar_y = ttkb.Scrollbar(alerts_frame, orient=VERTICAL, command=self.alerts_tree.yview, bootstyle="round")
        tree_scrollbar_y.pack(side=RIGHT, fill=Y)
        tree_scrollbar_x = ttkb.Scrollbar(alerts_frame, orient=HORIZONTAL, command=self.alerts_tree.xview, bootstyle="round")
        tree_scrollbar_x.pack(side=BOTTOM, fill=X)
        self.alerts_tree.pack(fill=BOTH, expand=True)
        self.alerts_tree.configure(yscrollcommand=tree_scrollbar_y.set, xscrollcommand=tree_scrollbar_x.set)
        self.alerts_tree.bind("<Double-1>", self.show_alerts_in_new_window)

        self.clear_results()
        
    def create_imagenes_widgets(self):
        frame = self.imagenes_tab
        frame.columnconfigure(0, weight=1)
        frame.columnconfigure(1, weight=3)
        frame.rowconfigure(0, weight=1)

        controls_frame = ttkb.Labelframe(frame, text="Acciones", padding=10)
        controls_frame.grid(row=0, column=0, sticky="nsew", padx=(0, 10))

        self.image_task_buttons = [
            ttkb.Button(controls_frame, text="AUTOMÁTICO\nBuscar y guardar pendientes", command=lambda: self.run_image_task(be.procesar_imagenes_automatico), bootstyle=PRIMARY),
            ttkb.Button(controls_frame, text="MANUAL\nSeleccionar archivo e imagen", command=lambda: self.run_image_task(be.procesar_una_imagen_manual), bootstyle=SECONDARY),
            ttkb.Button(controls_frame, text="Ver Registro", command=lambda: self.run_image_task(be.mostrar_registro_imagenes), bootstyle=INFO),
            ttkb.Button(controls_frame, text="Analizar Correspondencias", command=lambda: self.run_image_task(be.mostrar_correspondencias), bootstyle=INFO),
            ttkb.Button(controls_frame, text="Limpiar Registro", command=lambda: self.run_image_task(be.limpiar_registro_imagenes), bootstyle=DANGER)
        ]

        for btn in self.image_task_buttons:
            btn.pack(fill=X, pady=5, ipady=10 if "\n" in btn.cget("text") else 5)

        self.image_task_progressbar = ttkb.Progressbar(controls_frame, mode='indeterminate', bootstyle=PRIMARY)

        log_frame = ttkb.Labelframe(frame, text="Resultados y Registros", padding=10)
        log_frame.grid(row=0, column=1, sticky="nsew")
        self.image_log_text = scrolledtext.ScrolledText(log_frame, wrap=tk.WORD, state="disabled", font=("Courier New", 9))
        self.image_log_text.pack(fill=BOTH, expand=True)

    def create_config_widgets(self):
        frame = self.config_tab
        status_panel = ttkb.Labelframe(frame, text="Estado de Conexión", padding=15, bootstyle=INFO)
        status_panel.pack(fill=X, pady=(0, 10))
        self.main_status_indicator = ttkb.Label(status_panel, text="Verificando...", font=("-size 12 -weight bold"), anchor="center")
        self.main_status_indicator.pack(fill=X, ipady=5)
        folders_frame = ttkb.Labelframe(frame, text="Configuración de Carpetas", padding=15); folders_frame.pack(fill=X, pady=10)
        self.folder_vars = {}; folders = {"carpeta_excel": "Carpeta de Archivos Excel:","carpeta_fotos": "Carpeta de Capturas Temporales:","carpeta_liberaciones": "Carpeta Fotos de Resistencias:"}
        for key, text in folders.items():
            row_frame = ttkb.Frame(folders_frame); row_frame.pack(fill=X, pady=5)
            ttkb.Label(row_frame, text=text, width=30).pack(side=LEFT)
            var = tk.StringVar(); self.folder_vars[key] = var
            entry = ttkb.Entry(row_frame, textvariable=var); entry.pack(side=LEFT, fill=X, expand=True, padx=5)
            button = ttkb.Button(row_frame, text="Seleccionar...", command=lambda k=key, v=var: self.select_folder_for_config(k, v)); button.pack(side=LEFT)
        ttkb.Button(folders_frame, text="Guardar Configuración de Carpetas", command=self.save_folder_config, bootstyle=SUCCESS).pack(pady=15)
        whatsapp_frame = ttkb.Labelframe(frame, text="Acciones de WhatsApp", padding=15); whatsapp_frame.pack(fill=X, pady=10)
        ttkb.Button(whatsapp_frame, text="Verificar Estado de Conexión", command=self.update_whatsapp_status, bootstyle=INFO).pack(side=LEFT, padx=5, expand=True)
        ttkb.Button(whatsapp_frame, text="Abrir Panel de Green API", command=ws.abrir_panel_green_api).pack(side=LEFT, padx=5, expand=True)
        ttkb.Button(whatsapp_frame, text="Cerrar Sesión de WhatsApp", command=self.logout_whatsapp, bootstyle=DANGER).pack(side=LEFT, padx=5, expand=True)
        self.load_folder_config()

    def on_excel_file_select(self, event):
        self.clear_results()
        for widget in self.sample_buttons_container.winfo_children(): widget.destroy()
        self.sample_status_widgets.clear()

        selected_indices = self.excel_files_list.curselection()
        if not selected_indices: return
        
        self.selected_excel_file = self.excel_files_list.get(selected_indices[0])
        config = cm.cargar_configuracion()
        file_path = os.path.join(config["carpeta_excel"], self.selected_excel_file)
        
        resultado = be.obtener_muestras_disponibles(file_path)
        if not resultado["success"]:
            messagebox.showerror("Error al leer archivo", resultado["error"])
            return

        muestras = resultado.get("muestras", [])
        if not muestras:
            messagebox.showwarning("Sin Muestras", f"El archivo '{self.selected_excel_file}' no contiene hojas con formato 'MUESTRA A', etc.")
            return
        
        self.current_lote = "temp_lote"
        try:
            wb = openpyxl.load_workbook(file_path, data_only=True, read_only=True)
            hoja_a = wb[next(s.title for s in wb.worksheets if s.sheet_state == 'visible' and s.title.strip().upper().endswith("A"))]
            lote_val = hoja_a["F6"].value
            self.current_lote = ''.join(c for c in str(lote_val) if c.isdigit() or c == '-') if lote_val else "SIN_LOTE"
            wb.close()
        except Exception: pass
        
        for i, muestra in enumerate(muestras):
            is_sent = be.verificar_mensaje_duplicado(self.current_lote, muestra)
            style = "success" if is_sent else "primary-outline"
            
            button = ttkb.Button(self.sample_buttons_container, 
                                text=f"Muestra {muestra}", 
                                bootstyle=style, 
                                command=lambda s=muestra: self.run_analysis(sample_letter=s))
            # MEJORA: Colocar en una sola fila para scroll horizontal
            button.grid(row=0, column=i, padx=5, pady=5, sticky="ns")
            self.sample_status_widgets[muestra] = button
            self.sample_buttons_container.columnconfigure(i, weight=0) # No expandir columnas
        
        # MEJORA: Ajustar la altura del canvas al contenido
        self.root.update_idletasks()
        if muestras:
            height = self.sample_buttons_container.winfo_reqheight()
            self.sample_canvas.config(height=height)
        else:
            self.sample_canvas.config(height=1) # Colapsar si no hay muestras

        if not self.whatsapp_is_authorized:
            self.disable_buttons()
        else:
            self.enable_buttons()

    def run_analysis(self, sample_letter):
        if not self.selected_excel_file: return

        for btn in self.sample_status_widgets.values():
            btn.config(state="disabled")
        
        clicked_button = self.sample_status_widgets.get(sample_letter)
        if clicked_button:
            clicked_button.config(text="Analizando...", bootstyle=INFO)

        config = cm.cargar_configuracion()
        file_path = os.path.join(config["carpeta_excel"], self.selected_excel_file)
        
        if not os.path.exists(file_path):
            messagebox.showerror("Error de Archivo", f"El archivo {self.selected_excel_file} no se encuentra.")
            self.enable_buttons()
            return

        self.clear_results(keep_sample_buttons=True)
        self.set_status(f"Analizando {self.selected_excel_file}, Muestra {sample_letter}...")
        
        threading.Thread(target=self._analysis_thread, args=(file_path, sample_letter), daemon=True).start()

    def _analysis_thread(self, file_path, sheet_letter):
        try:
            result = be.analizar_archivo_para_gui(file_path, sheet_letter)
            if result["success"]:
                self.current_analysis_data = result["data"]
                self.root.after(0, self.update_analysis_results, result)
            else:
                self.root.after(0, lambda: messagebox.showerror("Error de Análisis", result["error"]))
                self.root.after(0, self.enable_buttons)
                return
            
            self.set_status("Generando vista previa de la imagen...")
            screenshot_result = be.tomar_captura_para_gui(file_path, result["data"]["hoja_nombre"], result["data"]["lote"])
            
            if screenshot_result["success"]:
                self.screenshot_path = screenshot_result["path"]
                self.root.after(0, lambda: self.display_thumbnail(self.screenshot_path, self.image_preview_label, size=(400, 500)))
            else:
                self.root.after(0, lambda: messagebox.showwarning("Error de Captura", screenshot_result["error"]))
        except Exception as e:
            detalle = str(e)  # se copia: "e" deja de existir al salir del except
            self.root.after(0, lambda detalle=detalle: messagebox.showerror("Error Crítico", f"Ocurrió un error inesperado: {detalle}"))
        finally:
            self.root.after(0, self.enable_buttons)
            self.update_whatsapp_status()

    def enable_buttons(self):
        if not self.whatsapp_is_authorized:
            self.disable_buttons()
            return
            
        lote = self.current_lote
        for muestra, button in self.sample_status_widgets.items():
            is_sent = be.verificar_mensaje_duplicado(lote, muestra) if lote else False
            style = "success" if is_sent else "primary-outline"
            button.config(state="normal", text=f"Muestra {muestra}", bootstyle=style)

        self.send_button.config(state="normal")
        
        if self.select_folder_button: self.select_folder_button.config(state="normal")
        self.excel_files_list.config(state="normal")
        self.set_status("Listo.")

    def disable_buttons(self):
        for button in self.sample_status_widgets.values(): button.config(state="disabled")
        self.send_button.config(state="disabled")
        if self.select_folder_button: self.select_folder_button.config(state="disabled")
        self.excel_files_list.config(state="disabled")
        self.set_status("🔴 WhatsApp no conectado. Conéctate para continuar.")
    
    def clear_results(self, keep_sample_buttons=False):
        if not keep_sample_buttons:
            for widget in self.sample_buttons_container.winfo_children(): widget.destroy()
            self.sample_status_widgets.clear()
            self.selected_excel_file = None
            self.current_lote = None

        self.message_text.delete('1.0', tk.END)
        self.image_preview_label.config(image='', text="La captura del Excel aparecerá aquí.")
        self.image_preview_label.photo_ref = None
        self.screenshot_path = None
        self.clear_additional_photo()
        self.residual_check_var.set(False)
        self.current_analysis_data = {}
        if self.sabor_alert_label: self.sabor_alert_label.pack_forget()
        
        if self.alerts_tree:
            for i in self.alerts_tree.get_children(): self.alerts_tree.delete(i)

    def _update_sample_status_ui(self, muestra, sent=True):
        if muestra in self.sample_status_widgets:
            button = self.sample_status_widgets[muestra]
            style = "success" if sent else "primary-outline"
            button.config(bootstyle=style)
            
    def send_to_whatsapp(self):
        message = self.message_text.get("1.0", tk.END).strip()
        if not message: messagebox.showwarning("Faltan Datos", "El mensaje está vacío."); return
        lote = self.current_analysis_data.get("lote", "SIN_LOTE"); muestra = self.current_analysis_data.get("muestra", "")
        if be.verificar_mensaje_duplicado(lote, muestra):
            if not messagebox.askyesno("Confirmar Reenvío", f"¡Atención!\n\nYa se envió una liberación para el LOTE {lote} - MUESTRA {muestra} el día de hoy.\n\n¿Estás seguro de que deseas enviarla de nuevo?"): return
        self.toggle_send_progress(True)
        threading.Thread(target=self._send_thread, args=(message,), daemon=True).start()

    def run_image_task(self, task_function):
        self.image_log_text.config(state="normal")
        self.image_log_text.delete('1.0', tk.END)
        self.image_log_text.config(state="disabled")
        
        self.set_status(f"Ejecutando: {task_function.__name__}...")
        
        for btn in self.image_task_buttons: btn.config(state="disabled")
        self.image_task_progressbar.pack(fill=X, pady=(10, 5))
        self.image_task_progressbar.start()

        threading.Thread(target=self._image_task_thread, args=(task_function,), daemon=True).start()

    def _image_task_thread(self, task_function):
        output_buffer = io.StringIO()
        try:
            with redirect_stdout(output_buffer): task_function()
        except Exception as e: print(f"\n\n--- ERROR CRÍTICO ---\n{e}")
        final_output = output_buffer.getvalue()
        self.root.after(0, self.finalize_image_task, final_output)

    def finalize_image_task(self, output):
        self.image_log_text.config(state="normal")
        self.image_log_text.insert('1.0', output)
        self.image_log_text.config(state="disabled")
        self.image_task_progressbar.stop()
        self.image_task_progressbar.pack_forget()
        for btn in self.image_task_buttons: btn.config(state="normal")
        self.set_status("Tarea de imágenes completada.")

    def handle_file_drop(self, event):
        filepath = self.root.tk.splitlist(event.data)[0]
        if filepath.lower().endswith(('.xlsx', '.xlsm')):
            filename = os.path.basename(filepath)
            if filename in self.excel_files_list.get(0, tk.END):
                idx = self.excel_files_list.get(0, tk.END).index(filename)
                self.excel_files_list.selection_clear(0, tk.END)
                self.excel_files_list.selection_set(idx)
                self.excel_files_list.activate(idx)
                self.excel_files_list.see(idx)
                self.on_excel_file_select(None)
            else: messagebox.showerror("Archivo no encontrado", f"El archivo '{filename}' no está en la carpeta de Excel configurada.")
        else: messagebox.showwarning("Archivo no válido", "Por favor, arrastre un archivo de Excel (.xlsx o .xlsm).")

    def handle_image_drop(self, event):
        filepath = self.root.tk.splitlist(event.data)[0]
        if filepath.lower().endswith(('.png', '.jpg', '.jpeg')):
            self.selected_additional_photo_path.set(filepath)
            self.additional_photo_label.config(text=os.path.basename(filepath))
            self.display_thumbnail(filepath, self.additional_photo_thumbnail)
            if self.delete_photo_button: self.delete_photo_button.config(state="normal")
        else: messagebox.showwarning("Archivo no válido", "Por favor, arrastre un archivo de imagen (.png, .jpg, .jpeg).")

    def show_image_in_new_window(self, image_path):
        if not image_path or not os.path.exists(image_path): return
        try:
            top = tk.Toplevel(self.root); top.title(os.path.basename(image_path))
            img = Image.open(image_path)
            img.thumbnail((self.root.winfo_screenwidth() * 0.8, self.root.winfo_screenheight() * 0.8), Image.Resampling.LANCZOS)
            photo = ImageTk.PhotoImage(img)
            label = ttk.Label(top, image=photo); label.image = photo; label.pack(padx=10, pady=10)
            top.transient(self.root); top.grab_set()
        except Exception as e: messagebox.showerror("Error de Imagen", f"No se pudo mostrar la imagen:\n{e}", parent=self.root)

    def display_thumbnail(self, image_path, label_widget, size=(200, 150)):
        try:
            img = Image.open(image_path); img.thumbnail(size, Image.Resampling.LANCZOS)
            photo_img = ImageTk.PhotoImage(img)
            label_widget.photo_ref = photo_img
            label_widget.config(image=photo_img, text="", cursor="hand2", relief="solid")
        except Exception as e: label_widget.config(image='', text=f"Error al cargar preview", cursor=""); label_widget.photo_ref = None

    def toggle_send_progress(self, show_progress=False):
        if show_progress:
            self.send_button.grid_forget()
            self.send_progressbar.grid(row=0, column=0, sticky="ew", ipady=5)
            self.send_progressbar.start()
        else:
            self.send_progressbar.stop()
            self.send_progressbar.grid_forget()
            self.send_button.grid(row=0, column=0, sticky="ew", ipady=5)

    def _update_whatsapp_status_thread(self):
        success, raw_state = ws.verificar_estado_green_api()
        if success and raw_state == "authorized": self.whatsapp_is_authorized = True; self.root.after(0, self.enable_buttons)
        else: self.whatsapp_is_authorized = False; self.root.after(0, self.disable_buttons)
        status_map = {"authorized": ("🟢 CONECTADO", "success"), "notAuthorized": ("🔴 NO AUTORIZADO", "danger"), "blocked": ("🚫 BLOQUEADO", "danger"), "sleepMode": ("🟡 EN SUSPENSIÓN", "warning"), "starting": ("🔵 INICIANDO", "info"), "connection_error": ("🔴 ERROR DE CONEXIÓN", "danger"), "error_http": ("🔴 ERROR DE API", "danger"), "unknown_error": ("⚪ DESCONOCIDO", "secondary")}
        status_text, style = status_map.get(raw_state, ("⚪ DESCONOCIDO", "secondary"))
        indicator_map = {"authorized": ("🟢", "success"),"notAuthorized": ("🔴", "danger"), "blocked": ("⚫", "danger"), "sleepMode": ("🟡", "warning"),"starting": ("🔵", "info"), "connection_error": ("🔴", "danger"), "error_http": ("🔴", "danger"),"unknown_error": ("⚪", "secondary")}
        indicator_text, indicator_style = indicator_map.get(raw_state, ("⚪", "secondary"))
        self.root.after(0, lambda: self.whatsapp_status_label.config(text=f"WhatsApp: {status_text}", bootstyle=style))
        if hasattr(self, 'main_status_indicator'): self.root.after(0, lambda: self.main_status_indicator.config(text=status_text, bootstyle=style))
        self.root.after(0, lambda: self.whatsapp_indicator.config(text=indicator_text, bootstyle=indicator_style))
            
    def _send_thread(self, message):
        try:
            # --- INICIO DE LA MODIFICACIÓN ---
            # Extraer todos los datos necesarios del diccionario del análisis actual
            lote = self.current_analysis_data.get("lote", "SIN_LOTE")
            muestra = self.current_analysis_data.get("muestra", "")
            proveedor = self.current_analysis_data.get("proveedor", "NO_PROVIDER") # <-- Dato adicional
            piscina = self.current_analysis_data.get("piscina", "NO_POOL")         # <-- Dato adicional
            # --- FIN DE LA MODIFICACIÓN ---

            additional_photo = self.selected_additional_photo_path.get()
            if muestra == "A" and not additional_photo:
                self.root.after(0, lambda: messagebox.showerror("Foto Requerida", "Para la MUESTRA A, es obligatorio adjuntar una foto de resistencia."))
                self.root.after(0, self.toggle_send_progress, False)
                return
            
            images_to_send = []
            if self.screenshot_path and os.path.exists(self.screenshot_path): 
                images_to_send.append(self.screenshot_path)
            
            if additional_photo and os.path.exists(additional_photo):
                if muestra == "A":
                    new_path, err = be.guardar_foto_adicional_muestra_a(additional_photo, lote)
                    if err: self.root.after(0, lambda: messagebox.showwarning("Error al Guardar Foto", err))
                    images_to_send.append(new_path if new_path else additional_photo)
                else: 
                    images_to_send.append(additional_photo)
            
            if not images_to_send:
                self.root.after(0, lambda: messagebox.showwarning("Faltan Datos", "No hay ninguna imagen para enviar."));
                self.root.after(0, self.toggle_send_progress, False)
                return
            
            success, details = ws.enviar_mensaje_con_imagenes_gui(lote, muestra, message, images_to_send)
            
            if success:
                # --- MODIFICACIÓN EN LA LLAMADA A LA FUNCIÓN ---
                # Ahora pasamos el proveedor y la piscina al guardar el registro
                be.guardar_mensaje_enviado(lote, muestra, proveedor, piscina)
                # --- FIN DE LA MODIFICACIÓN ---

                self.root.after(0, self._update_sample_status_ui, muestra, True)
                self.root.after(0, lambda: messagebox.showinfo("Envío Exitoso", f"Mensaje y {details['enviados']}/{details['totales']} imágenes enviadas."))
                self.root.after(0, self.clear_results)
            else: 
                self.root.after(0, lambda: messagebox.showerror("Error de Envío", f"No se pudo enviar el mensaje.\n\nDetalle: {details['error']}"))
        
        except Exception as e:
            detalle = str(e)  # se copia: "e" deja de existir al salir del except
            self.root.after(0, lambda detalle=detalle: messagebox.showerror("Error Crítico de Envío", f"Ocurrió un error inesperado: {detalle}"))
        finally:
            self.root.after(0, self.toggle_send_progress, False)
            self.update_whatsapp_status()

    def set_status(self, text):
        self.status_label.config(text=text)

    def update_whatsapp_status(self):
        self.set_status("Verificando estado de WhatsApp..."); threading.Thread(target=self._update_whatsapp_status_thread, daemon=True).start()
    
    def update_analysis_results(self, result):
        for i in self.alerts_tree.get_children(): self.alerts_tree.delete(i)
        
        local_alerts = result["alerts"].get("local", [])
        global_alerts = result["alerts"].get("global", [])
        all_alerts = [("Local", a) for a in local_alerts] + [("Global", a) for a in global_alerts]

        if not all_alerts:
            self.alerts_tree.insert("", "end", values=("", "✅", "No se encontraron alertas"))
        else:
            for tipo, desc in all_alerts:
                prioridad = "Error" if "error" in desc.lower() else "Advertencia"
                self.alerts_tree.insert("", "end", values=(tipo, prioridad, desc))

        if "REMUESTREADO POR" in result["message"]: self.sabor_alert_label.pack(fill=X, pady=(15, 5), ipady=5)
        else: self.sabor_alert_label.pack_forget()
            
        self.message_text.delete('1.0', tk.END)
        self.message_text.insert('1.0', result["message"])

        if local_alerts:
            popup_message = f"Se encontraron {len(local_alerts)} advertencia(s) en esta muestra.\n\n" + "\n".join(f"● {a}" for a in local_alerts) + "\n\nPor favor, revísalas cuidadosamente antes de continuar."
            self.show_warning_popup(popup_message)
    
    def show_warning_popup(self, message):
        alert_window = tk.Toplevel(self.root)
        alert_window.title("¡Advertencias de Validación!")
        alert_window.geometry("550x350")
        alert_window.transient(self.root); alert_window.grab_set()
        
        main_frame = ttkb.Frame(alert_window, padding=20)
        main_frame.pack(fill=BOTH, expand=True)
        
        ttkb.Label(main_frame, text="¡Atención!", font=("Segoe UI", 16, "bold"), bootstyle=WARNING).pack(pady=(0, 10))
        
        text_widget = scrolledtext.ScrolledText(main_frame, wrap=tk.WORD, height=8, font=("Segoe UI", 10))
        text_widget.pack(fill=BOTH, expand=True, pady=5)
        text_widget.insert('1.0', message)
        text_widget.config(state="disabled")
        
        ttkb.Button(main_frame, text="Entendido, Continuar", command=alert_window.destroy, bootstyle=WARNING).pack(pady=(10, 0), ipady=5)
        
        self.root.wait_window(alert_window)

    def on_residual_check_toggle(self):
        texto_residual = "\n**RESIDUAL EN ANÁLISIS**"
        contenido_actual = self.message_text.get("1.0", tk.END)
        self.message_text.delete("1.0", tk.END)
        if self.residual_check_var.get(): self.message_text.insert("1.0", contenido_actual.strip() + texto_residual)
        else: self.message_text.insert("1.0", contenido_actual.replace(texto_residual, "").strip())

    def select_additional_photo(self):
        config = cm.cargar_configuracion()
        filepath = filedialog.askopenfilename(initialdir=config.get("carpeta_liberaciones", os.path.expanduser("~")),title="Selecciona una imagen",filetypes=[("Imágenes", "*.png;*.jpg;*.jpeg")])
        if filepath:
            self.selected_additional_photo_path.set(filepath)
            self.additional_photo_label.config(text=os.path.basename(filepath))
            self.display_thumbnail(filepath, self.additional_photo_thumbnail)
            if self.delete_photo_button: self.delete_photo_button.config(state="normal")
    
    def check_initial_config(self):
        config_ok, message = cm.verificar_configuracion()
        if not config_ok: messagebox.showwarning("Configuración Requerida", f"La configuración no está completa.\n\n{message}\n\nPor favor, ve a la pestaña 'Configuración'."); self.notebook.select(self.config_tab)
        else: config = cm.cargar_configuracion(); self.load_excel_files(config["carpeta_excel"])

    def select_excel_folder(self):
        config = cm.cargar_configuracion(); folder = filedialog.askdirectory(initialdir=config.get("carpeta_excel", os.path.expanduser("~")))
        if folder: config["carpeta_excel"] = folder; cm.guardar_configuracion(config); self.load_excel_files(folder); self.set_status(f"Carpeta seleccionada: {folder}")

    def load_excel_files(self, folder_path):
        self.excel_files_list.delete(0, tk.END)
        try:
            files = sorted([f for f in os.listdir(folder_path) if f.endswith((".xlsx", ".xlsm")) and not f.startswith("~")])
            for file in files: self.excel_files_list.insert(tk.END, file)
        except Exception as e: messagebox.showerror("Error", f"No se pudo leer la carpeta: {e}")

    def select_folder_for_config(self, key, var):
        current_path = var.get();
        if not os.path.isdir(current_path): current_path = os.path.expanduser("~")
        folder = filedialog.askdirectory(initialdir=current_path)
        if folder: var.set(folder)

    def load_folder_config(self):
        config = cm.cargar_configuracion()
        for key, var in self.folder_vars.items(): var.set(config.get(key, ""))

    def save_folder_config(self):
        config = cm.cargar_configuracion()
        for key, var in self.folder_vars.items(): config[key] = var.get()
        for key in ["carpeta_fotos", "carpeta_liberaciones"]:
            if config[key]: os.makedirs(config[key], exist_ok=True)
        config["primera_configuracion"] = False
        if cm.guardar_configuracion(config): messagebox.showinfo("Configuración", "La configuración se guardó correctamente."); self.load_excel_files(config["carpeta_excel"])
        else: messagebox.showerror("Error", "No se pudo guardar la configuración.")

    def logout_whatsapp(self):
        if messagebox.askyesno("Confirmar Cierre de Sesión", "¿Estás seguro de que deseas cerrar la sesión de WhatsApp?"):
            self.disable_buttons()
            self.set_status("Cerrando sesión de WhatsApp...")
            threading.Thread(target=self._logout_thread, daemon=True).start()

    def _logout_thread(self):
        success, message = ws.cerrar_sesion_whatsapp_gui()
        if success: self.root.after(0, lambda: messagebox.showinfo("Sesión Cerrada", message))
        else: self.root.after(0, lambda: messagebox.showerror("Error", message))
        self.root.after(0, self.update_whatsapp_status)

    def clear_additional_photo(self):
        """Limpia la foto adicional seleccionada y resetea los widgets."""
        self.selected_additional_photo_path.set("")
        self.additional_photo_label.config(text="Ninguna foto seleccionada.")
        self.additional_photo_thumbnail.config(image='', text="⬇  Arrastre una imagen aquí", cursor="", relief="solid")
        self.additional_photo_thumbnail.photo_ref = None
        if self.delete_photo_button:
            self.delete_photo_button.config(state="disabled")

    def show_alerts_in_new_window(self, event):
        """Muestra las alertas en una ventana Toplevel más grande al hacer doble clic."""
        if not self.alerts_tree.get_children():
            return

        top = tk.Toplevel(self.root)
        top.title("Detalle de Advertencias de Validación")
        top.geometry("800x400")
        top.transient(self.root)
        top.grab_set()

        frame = ttkb.Frame(top, padding=10)
        frame.pack(fill=BOTH, expand=True)

        cols = ("tipo", "prioridad", "descripcion")
        tree = ttk.Treeview(frame, columns=cols, show='headings', selectmode="none")
        
        tree.heading("tipo", text="Tipo")
        tree.heading("prioridad", text="Prioridad")
        tree.heading("descripcion", text="Descripción")
        
        tree.column("tipo", width=80, stretch=False, anchor="center")
        tree.column("prioridad", width=100, stretch=False, anchor="center")
        tree.column("descripcion", width=550)

        # Copiar datos desde el Treeview original
        for item_id in self.alerts_tree.get_children():
            values = self.alerts_tree.item(item_id, 'values')
            tree.insert("", "end", values=values)
            
        # Añadir scrollbars a la nueva ventana
        scrollbar_y = ttkb.Scrollbar(frame, orient=VERTICAL, command=tree.yview, bootstyle="round")
        scrollbar_y.pack(side=RIGHT, fill=Y)
        scrollbar_x = ttkb.Scrollbar(frame, orient=HORIZONTAL, command=tree.xview, bootstyle="round")
        scrollbar_x.pack(side=BOTTOM, fill=X)
        tree.pack(fill=BOTH, expand=True)
        tree.configure(yscrollcommand=scrollbar_y.set, xscrollcommand=scrollbar_x.set)