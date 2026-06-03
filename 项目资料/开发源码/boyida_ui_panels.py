from tkinter import Canvas, Text, ttk


def build_left_panel(app, workbench, compact: bool, ultra_compact: bool) -> None:
    del ultra_compact
    panel_pad = (8 if compact else 10, 10 if compact else 12, 8 if compact else 10, 6 if compact else 8)
    status_gap = (4, 5 if compact else 7)
    path_gap = (0, 10 if compact else 12)
    button_gap = 8 if compact else 9
    primary_gap = (0, button_gap)
    action_gap = (0, button_gap)
    update_gap = (0, 0)
    batch_ipady = 3 if compact else 4
    primary_ipady = batch_ipady
    left_panel = ttk.LabelFrame(
        workbench,
        text="批次与操作",
        style="Panel.TLabelframe",
        padding=panel_pad,
    )
    left_panel.grid(row=0, column=0, sticky="nsew", padx=(0, 8 if compact else 10))
    left_panel.columnconfigure(0, weight=1)
    left_panel.rowconfigure(0, weight=0)

    left_content = ttk.Frame(left_panel, style="Panel.TFrame")
    left_content.grid(row=0, column=0, sticky="new", pady=(6 if compact else 8, 0))
    left_content.columnconfigure(0, weight=1)
    ttk.Label(left_content, textvariable=app.input_status, style="BatchStatus.TLabel").grid(row=0, column=0, sticky="ew", pady=status_gap)
    ttk.Label(left_content, textvariable=app.output_status, style="BatchPath.TLabel").grid(row=1, column=0, sticky="ew", pady=path_gap)
    ttk.Button(left_content, text="导入图片", command=app._choose_image_dir, style="Batch.TButton").grid(row=2, column=0, sticky="ew", pady=(0, button_gap), ipady=batch_ipady)
    ttk.Button(left_content, text="一键出表", command=app.run_all, style="BatchPrimary.TButton").grid(
        row=3,
        column=0,
        sticky="ew",
        pady=primary_gap,
        ipady=primary_ipady,
    )

    action_grid = ttk.Frame(left_content, style="Panel.TFrame")
    action_grid.grid(row=4, column=0, sticky="ew", pady=(0, button_gap))
    action_grid.columnconfigure(0, weight=1)
    action_grid.columnconfigure(1, weight=1)
    action_buttons = [
        ("检测环境", app.check_ocr_env),
        ("打开结果", app.open_work_dir),
        ("OCR 设置", app.open_ocr_settings),
        ("GPS 设置", app.open_gps_settings),
    ]
    for index, (text, command) in enumerate(action_buttons):
        row = index // 2
        column = index % 2
        ttk.Button(action_grid, text=text, command=command, style="Batch.TButton").grid(
            row=row,
            column=column,
                sticky="ew",
                padx=(0, 4) if column == 0 else (4, 0),
                pady=(0, button_gap) if row == 0 else (0, 0),
                ipady=batch_ipady,
            )
    ttk.Button(left_content, text="检查更新", command=app.check_for_updates, style="Batch.TButton").grid(
        row=5,
        column=0,
        sticky="ew",
        pady=update_gap,
        ipady=batch_ipady,
    )


def build_center_panel(app, workbench, compact: bool, ultra_compact: bool, theme: dict) -> None:
    center_panel = ttk.LabelFrame(workbench, text="流程管线", style="Panel.TLabelframe", padding=(8 if compact else 10, 8 if compact else 10))
    center_panel.grid(row=0, column=1, sticky="nsew", padx=(0, 8 if compact else 10))
    center_panel.rowconfigure(0, weight=0)
    center_panel.rowconfigure(1, weight=1)
    center_panel.columnconfigure(0, weight=1)

    pipeline_width = 300 if ultra_compact else (390 if compact else 470)
    pipeline_height = 116 if ultra_compact else (122 if compact else 132)
    app.pipeline_canvas = Canvas(center_panel, width=pipeline_width, height=pipeline_height, bg=theme["log"], highlightthickness=0)
    app.pipeline_canvas.grid(row=0, column=0, sticky="ew")
    app.pipeline_canvas.bind("<Configure>", lambda _event: app._draw_pipeline_panel(app.pipeline_canvas))
    app._draw_pipeline_panel(app.pipeline_canvas)

    progress_panel = ttk.Frame(center_panel, style="Panel.TFrame")
    progress_panel.grid(row=1, column=0, sticky="nsew", pady=(10, 0))
    progress_panel.columnconfigure(0, weight=1)
    progress_panel.rowconfigure(2, weight=1)
    progress_header = ttk.Frame(progress_panel, style="Panel.TFrame")
    progress_header.grid(row=0, column=0, sticky="ew", pady=(0, 5))
    progress_header.columnconfigure(0, weight=1)
    ttk.Label(progress_header, textvariable=app.progress_title, style="Panel.TLabel").grid(row=0, column=0, sticky="w")
    ttk.Label(progress_header, textvariable=app.progress_detail, style="PanelMuted.TLabel").grid(row=0, column=1, sticky="e", padx=(12, 8))
    ttk.Label(progress_header, textvariable=app.progress_percent, style="ProgressPercent.TLabel").grid(row=0, column=2, sticky="e")
    ttk.Progressbar(
        progress_panel,
        variable=app.progress_value,
        maximum=100,
        mode="determinate",
        style="Ops.Horizontal.TProgressbar",
    ).grid(row=1, column=0, sticky="ew", pady=(0, 6))

    progress_log_frame = ttk.Frame(progress_panel, style="Panel.TFrame")
    progress_log_frame.grid(row=2, column=0, sticky="nsew")
    progress_log_frame.columnconfigure(0, weight=1)
    progress_log_frame.rowconfigure(0, weight=1)
    app.progress_log_text = Text(
        progress_log_frame,
        wrap="word",
        height=4 if compact else 5,
        bg=theme["log"],
        fg=theme["muted"],
        insertbackground=theme["cyan"],
        selectbackground=theme["selected"],
        relief="flat",
        borderwidth=0,
        padx=8,
        pady=6,
        state="disabled",
        font=("Consolas", 10),
    )
    app.progress_log_text.tag_configure("section", foreground=theme["cyan"], font=("Microsoft YaHei UI", 10, "bold"))
    app.progress_log_text.tag_configure("success", foreground=theme["green"])
    app.progress_log_text.tag_configure("error", foreground="#ffd1d8", background="#3a0d16")
    progress_scroll = ttk.Scrollbar(progress_log_frame, orient="vertical", command=app.progress_log_text.yview)
    app.progress_log_text.configure(yscrollcommand=progress_scroll.set)
    app.progress_log_text.grid(row=0, column=0, sticky="nsew")
    progress_scroll.grid(row=0, column=1, sticky="ns")


def build_right_panel(app, workbench, compact: bool, ultra_compact: bool, theme: dict, ocr_engine_labels: dict, ocr_profile_labels: dict) -> None:
    right_panel = ttk.Frame(workbench, style="App.TFrame")
    right_panel.grid(row=0, column=2, sticky="new")
    right_panel.columnconfigure(0, weight=1)
    right_panel.rowconfigure(1, weight=0)

    telemetry = ttk.LabelFrame(right_panel, text="车队遥测", style="Panel.TLabelframe", padding=(8, 8))
    telemetry.grid(row=0, column=0, sticky="ew")
    route_width = 260 if ultra_compact else (288 if compact else 310)
    route_height = 112 if compact else 118
    app.route_canvas = Canvas(telemetry, width=route_width, height=route_height, bg=theme["log"], highlightthickness=0, cursor="hand2")
    app.route_canvas.pack(fill="x")
    app.route_canvas.bind("<Button-1>", lambda _event: app.refresh_gps())
    app.route_canvas.bind("<Configure>", lambda _event: app._draw_route_panel(app.route_canvas))
    app._draw_route_panel(app.route_canvas)

    ocr_panel = ttk.LabelFrame(right_panel, text="OCR 控制", style="Panel.TLabelframe", padding=(8 if compact else 9, 6 if compact else 7))
    ocr_panel.grid(row=1, column=0, sticky="ew", pady=(5 if compact else 6, 0))
    ocr_panel.columnconfigure(0, weight=1)
    ttk.Label(ocr_panel, text="OCR 引擎", style="PanelMuted.TLabel").grid(row=0, column=0, sticky="w")
    engine_combo = ttk.Combobox(ocr_panel, textvariable=app.ocr_engine, values=list(ocr_engine_labels.values()), state="readonly", width=30)
    engine_combo.grid(row=1, column=0, sticky="ew", pady=(2, 5))
    ttk.Label(ocr_panel, text="OCR 档位 / 实际模型", style="PanelMuted.TLabel").grid(row=2, column=0, sticky="w")
    combo = ttk.Combobox(ocr_panel, textvariable=app.ocr_profile, values=list(ocr_profile_labels.values()), state="readonly", width=38)
    combo.grid(row=3, column=0, sticky="ew", pady=(2, 3))
    ttk.Checkbutton(ocr_panel, text="已有 OCR JSON 时跳过整批 OCR", variable=app.skip_ocr_if_json).grid(row=4, column=0, sticky="w", pady=(0, 6))
    ttk.Label(ocr_panel, text="ZHIPU API Key", style="PanelMuted.TLabel").grid(row=5, column=0, sticky="w")
    ttk.Entry(ocr_panel, textvariable=app.api_key, show="*").grid(row=6, column=0, sticky="ew", pady=(2, 4))
    ttk.Checkbutton(ocr_panel, text="记住 Key（仅本机）", variable=app.remember_api_key).grid(row=7, column=0, sticky="w")
