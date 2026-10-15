using System.Diagnostics;
using System.Text;
using MetaCleaner.Core;

namespace MetaCleaner.App;

sealed class FileEntry
{
    public string Path = "";
    public string Root = "";
    public DocKind Kind;
    public DocProps Original = new();
    public DocProps Edited = new();
    public InspectResult Result = new();
    public bool Dirty => Fields.All.Any(f => Edited.Get(f.Key) != Original.Get(f.Key));
}

public sealed class MainForm : Form
{
    static readonly Color ColEdited = Color.FromArgb(255, 241, 168);
    static readonly Color ColSuspicious = Color.FromArgb(255, 212, 212);
    static readonly Color ColHigh = Color.FromArgb(200, 40, 40), ColMed = Color.FromArgb(210, 130, 0), ColOk = Color.FromArgb(30, 140, 60);

    static readonly string[] BulkKeys =
    {
        Fields.Creator, Fields.LastModifiedBy, Fields.Description, Fields.Title, Fields.Subject,
        Fields.Keywords, Fields.Category, Fields.ContentStatus, Fields.Company, Fields.Manager,
    };

    readonly List<FileEntry> entries = new();
    readonly Dictionary<string, (CheckBox Chk, TextBox Box)> bulk = new();
    readonly Dictionary<string, CheckBox> optionChecks = new();
    List<Profile> profiles = ProfileStore.Load();
    bool filling, busy;

    readonly DataGridView gridProps, gridFind;
    readonly TextBox logBox;
    readonly TabControl tabs;
    readonly CheckBox chkRecursive;
    readonly ToolStripProgressBar progress;
    readonly ToolStripStatusLabel status;
    readonly ToolStrip toolbar;
    readonly SplitContainer split;
    readonly FlowLayoutPanel side;

    CheckBox chkApp = null!, chkPlausible = null!, chkBackup = null!, chkOffice = null!;
    RadioButton rbDateKeep = null!, rbDateSet = null!, rbDateRand = null!, rbOutFolder = null!, rbInPlace = null!;
    DateTimePicker dtCreated = null!, dtModified = null!, dtFrom = null!, dtTo = null!;
    TextBox txtReviewer = null!;
    ComboBox cmbProfile = null!, cmbDash = null!;

    public MainForm()
    {
        Text = "MetaCleaner — свойства и следы в документах Office";
        Width = 1450; Height = 880;
        MinimumSize = new Size(1000, 600);
        StartPosition = FormStartPosition.CenterScreen;
        Font = new Font("Segoe UI", 9f);
        AllowDrop = true;
        DragEnter += (_, e) => { if (e.Data?.GetDataPresent(DataFormats.FileDrop) == true) e.Effect = DragDropEffects.Copy; };
        DragDrop += (_, e) => { if (e.Data?.GetData(DataFormats.FileDrop) is string[] p) AddPaths(p); };

        // ---------- панель инструментов ----------
        toolbar = new ToolStrip { GripStyle = ToolStripGripStyle.Hidden, Padding = new Padding(6, 3, 6, 3) };
        toolbar.Items.Add(Btn("📂  Открыть папку", OpenFolder));
        toolbar.Items.Add(Btn("📄  Добавить файлы", AddFilesDialog));
        chkRecursive = new CheckBox { Text = "Включая подпапки", Checked = true, BackColor = Color.Transparent, AutoSize = true };
        toolbar.Items.Add(new ToolStripControlHost(chkRecursive) { Margin = new Padding(8, 2, 8, 0) });
        toolbar.Items.Add(new ToolStripSeparator());
        toolbar.Items.Add(Btn("🔄  Перечитать", Reload));
        toolbar.Items.Add(Btn("✖  Очистить список", () => { if (ConfirmDiscard()) { entries.Clear(); RefreshAll(); } }));
        toolbar.Items.Add(new ToolStripSeparator());
        toolbar.Items.Add(Btn("📊  Экспорт в CSV", ExportCsv));
        var saveBtn = Btn("💾  Сохранить изменения", SaveAll);
        saveBtn.Font = new Font(Font, FontStyle.Bold);
        saveBtn.Alignment = ToolStripItemAlignment.Right;
        toolbar.Items.Add(saveBtn);

        // ---------- строка состояния ----------
        var statusStrip = new StatusStrip();
        status = new ToolStripStatusLabel("Откройте папку или перетащите файлы .docx / .pptx / .xlsx в окно")
            { Spring = true, TextAlign = ContentAlignment.MiddleLeft };
        progress = new ToolStripProgressBar { Width = 220, Visible = false };
        statusStrip.Items.Add(status);
        statusStrip.Items.Add(progress);

        // ---------- вкладки ----------
        tabs = new TabControl { Dock = DockStyle.Fill };
        gridProps = NewGrid(editable: true);
        gridFind = NewGrid(editable: false);
        logBox = new TextBox
        {
            Dock = DockStyle.Fill, Multiline = true, ReadOnly = true, ScrollBars = ScrollBars.Both,
            WordWrap = false, Font = new Font("Consolas", 9f), BackColor = SystemColors.Window,
        };

        var tabProps = new TabPage("Свойства");
        var legend = new Label
        {
            Dock = DockStyle.Top, Height = 26, Padding = new Padding(6, 5, 0, 0), ForeColor = Color.DimGray,
            Text = "Ячейки можно править вручную.  Жёлтым — изменено, но ещё не сохранено.  " +
                   "Красным — подозрительное значение.  Выделите строки и примените массовую замену справа.",
        };
        tabProps.Controls.Add(gridProps);
        tabProps.Controls.Add(legend);
        var tabFind = new TabPage("Проверка");
        tabFind.Controls.Add(gridFind);
        var tabLog = new TabPage("Журнал");
        tabLog.Controls.Add(logBox);
        tabs.TabPages.AddRange(new[] { tabProps, tabFind, tabLog });

        BuildPropsGrid();
        BuildFindGrid();

        // ---------- правая панель ----------
        side = new FlowLayoutPanel
        {
            Dock = DockStyle.Fill, FlowDirection = FlowDirection.TopDown, WrapContents = false,
            AutoScroll = true, Padding = new Padding(6),
        };
        BuildSidePanel();

        split = new SplitContainer { Dock = DockStyle.Fill, FixedPanel = FixedPanel.Panel2, SplitterWidth = 6 };
        split.Panel1.Controls.Add(tabs);
        split.Panel2.Controls.Add(side);

        Controls.Add(split);
        Controls.Add(toolbar);
        Controls.Add(statusStrip);

        Load += (_, _) =>
        {
            try { split.SplitterDistance = Math.Max(400, split.Width - 400); } catch { }
            Log("MetaCleaner запущен. " + (OfficeSession.AnyAvailable()
                ? "Microsoft Office найден — доступно пересохранение через Office."
                : "Microsoft Office не найден — пересохранение через Office недоступно."));
        };
        FormClosing += (_, e) =>
        {
            if (busy) { e.Cancel = true; return; }
            if (!ConfirmDiscard()) e.Cancel = true;
        };
    }

    // ======================= построение интерфейса =======================

    ToolStripButton Btn(string text, Action a)
    {
        var b = new ToolStripButton(text) { DisplayStyle = ToolStripItemDisplayStyle.Text, Margin = new Padding(2, 1, 2, 1) };
        b.Click += (_, _) => { if (!busy) a(); };
        return b;
    }

    static DataGridView NewGrid(bool editable)
    {
        var g = new DataGridView
        {
            Dock = DockStyle.Fill, AllowUserToAddRows = false, AllowUserToDeleteRows = false,
            AllowUserToResizeRows = false, RowHeadersWidth = 28, BackgroundColor = SystemColors.Window,
            BorderStyle = BorderStyle.None, ReadOnly = !editable,
            SelectionMode = editable ? DataGridViewSelectionMode.CellSelect : DataGridViewSelectionMode.FullRowSelect,
            EditMode = DataGridViewEditMode.EditOnKeystrokeOrF2, ColumnHeadersHeightSizeMode = DataGridViewColumnHeadersHeightSizeMode.AutoSize,
            EnableHeadersVisualStyles = true, ShowCellToolTips = true,
        };
        g.DefaultCellStyle.SelectionBackColor = Color.FromArgb(204, 228, 247);
        g.DefaultCellStyle.SelectionForeColor = Color.Black;
        g.RowTemplate.Height = 24;
        typeof(DataGridView).GetProperty("DoubleBuffered",
            System.Reflection.BindingFlags.Instance | System.Reflection.BindingFlags.NonPublic)?.SetValue(g, true);
        return g;
    }

    static DataGridViewTextBoxColumn Col(DataGridView g, string name, string header, int width, bool ro)
    {
        var c = new DataGridViewTextBoxColumn
        {
            Name = name, HeaderText = header, Width = width, ReadOnly = ro, SortMode = DataGridViewColumnSortMode.Automatic,
        };
        g.Columns.Add(c);
        return c;
    }

    void BuildPropsGrid()
    {
        Col(gridProps, "file", "Файл", 230, true).Frozen = true;
        Col(gridProps, "kind", "Тип", 80, true);
        Col(gridProps, "score", "Оценка", 85, true);
        foreach (var (k, label) in Fields.All)
            Col(gridProps, k, label, k == Fields.Description ? 220 : k is Fields.Created or Fields.Modified ? 125 : 130, false);
        Col(gridProps, "custom", "Польз. свойства", 160, true);
        Col(gridProps, "count", "Следов", 60, true);

        gridProps.CellFormatting += (_, e) =>
        {
            if (e.RowIndex < 0 || e.CellStyle == null || gridProps.Rows[e.RowIndex].Tag is not FileEntry en) return;
            var key = gridProps.Columns[e.ColumnIndex].Name;
            if (key == "score")
            {
                e.CellStyle.ForeColor = SevColor(en.Result.Score);
                e.CellStyle.Font = new Font(Font, FontStyle.Bold);
                return;
            }
            if (Fields.All.All(f => f.Key != key)) return;
            if (en.Edited.Get(key) != en.Original.Get(key)) e.CellStyle.BackColor = ColEdited;
            else if (en.Result.SuspiciousFields.Contains(key)) e.CellStyle.BackColor = ColSuspicious;
        };
        gridProps.CellValueChanged += (_, e) =>
        {
            if (filling || e.RowIndex < 0 || gridProps.Rows[e.RowIndex].Tag is not FileEntry en) return;
            var key = gridProps.Columns[e.ColumnIndex].Name;
            if (Fields.All.All(f => f.Key != key)) return;
            en.Edited[key] = gridProps.Rows[e.RowIndex].Cells[e.ColumnIndex].Value?.ToString()?.Trim() ?? "";
            gridProps.InvalidateRow(e.RowIndex);
            UpdateStatus();
        };
        gridProps.KeyDown += (_, e) =>
        {
            // Delete — очистить выделенные ячейки свойств
            if (e.KeyCode != Keys.Delete || gridProps.IsCurrentCellInEditMode) return;
            foreach (DataGridViewCell c in gridProps.SelectedCells)
                if (!c.ReadOnly) c.Value = "";
            e.Handled = true;
        };
    }

    void BuildFindGrid()
    {
        Col(gridFind, "file", "Файл", 230, true);
        Col(gridFind, "sev", "Уровень", 85, true);
        Col(gridFind, "cat", "Категория", 110, true);
        var msg = Col(gridFind, "msg", "Что найдено", 600, true);
        msg.AutoSizeMode = DataGridViewAutoSizeColumnMode.Fill;
        Col(gridFind, "fix", "Как исправить", 170, true);
        gridFind.CellFormatting += (_, e) =>
        {
            if (e.RowIndex < 0 || e.CellStyle == null || gridFind.Rows[e.RowIndex].Tag is not Finding f) return;
            if (gridFind.Columns[e.ColumnIndex].Name == "sev")
            {
                e.CellStyle.ForeColor = SevColor(f.Severity);
                e.CellStyle.Font = new Font(Font, FontStyle.Bold);
            }
        };
        gridFind.CellDoubleClick += (_, e) =>
        {
            // двойной щелчок — перейти к строке файла на вкладке «Свойства»
            if (e.RowIndex < 0) return;
            var path = gridFind.Rows[e.RowIndex].Cells["file"].ToolTipText;
            foreach (DataGridViewRow r in gridProps.Rows)
                if (r.Tag is FileEntry en && en.Path == path)
                {
                    tabs.SelectedIndex = 0;
                    gridProps.ClearSelection();
                    gridProps.CurrentCell = r.Cells["file"];
                    r.Selected = true;
                    break;
                }
        };
    }

    GroupBox Group(string title, Control content)
    {
        var g = new GroupBox { Text = title, Width = 370, AutoSize = true, AutoSizeMode = AutoSizeMode.GrowAndShrink, Padding = new Padding(8, 6, 8, 8) };
        content.Location = new Point(8, 22);
        g.Controls.Add(content);
        return g;
    }

    static TableLayoutPanel Table(int cols)
    {
        var t = new TableLayoutPanel { ColumnCount = cols, AutoSize = true, AutoSizeMode = AutoSizeMode.GrowAndShrink, Width = 350 };
        return t;
    }

    CheckBox Opt(string key, string text, bool on, string? tip = null)
    {
        var c = new CheckBox { Text = text, Checked = on, AutoSize = true, Margin = new Padding(3, 3, 3, 1) };
        if (tip != null) new ToolTip { AutoPopDelay = 15000 }.SetToolTip(c, tip);
        optionChecks[key] = c;
        return c;
    }

    static DateTimePicker Dtp(DateTime value) => new()
    {
        Format = DateTimePickerFormat.Custom, CustomFormat = "dd.MM.yyyy HH:mm", Width = 140, Value = value,
    };

    void BuildSidePanel()
    {
        // --- профиль ---
        var prof = Table(3);
        prof.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 100));
        prof.ColumnStyles.Add(new ColumnStyle(SizeType.AutoSize));
        prof.ColumnStyles.Add(new ColumnStyle(SizeType.AutoSize));
        cmbProfile = new ComboBox { DropDownStyle = ComboBoxStyle.DropDownList, Dock = DockStyle.Fill };
        cmbProfile.SelectedIndexChanged += (_, _) =>
        {
            if (cmbProfile.SelectedIndex > 0) ApplyProfile(profiles[cmbProfile.SelectedIndex - 1]);
        };
        var bSave = new Button { Text = "Сохранить", AutoSize = true };
        bSave.Click += (_, _) => SaveProfile();
        var bDel = new Button { Text = "Удалить", AutoSize = true };
        bDel.Click += (_, _) => DeleteProfile();
        prof.Controls.Add(cmbProfile, 0, 0);
        prof.Controls.Add(bSave, 1, 0);
        prof.Controls.Add(bDel, 2, 0);
        FillProfiles();
        side.Controls.Add(Group("Профиль", prof));

        // --- массовая замена ---
        var t = Table(2);
        t.ColumnStyles.Add(new ColumnStyle(SizeType.Absolute, 135));
        t.ColumnStyles.Add(new ColumnStyle(SizeType.Absolute, 205));
        int row = 0;
        var hint = new Label
        {
            Text = "Отметьте поля, которые нужно заменить. Пустое отмеченное поле = очистить.",
            AutoSize = true, MaximumSize = new Size(340, 0), ForeColor = Color.DimGray, Margin = new Padding(3, 0, 3, 6),
        };
        t.Controls.Add(hint, 0, row);
        t.SetColumnSpan(hint, 2);
        row++;
        foreach (var k in BulkKeys)
        {
            var c = new CheckBox { Text = Fields.Label(k), AutoSize = true, Margin = new Padding(3, 5, 3, 0) };
            var b = new TextBox { Width = 200 };
            if (k == Fields.LastModifiedBy) b.PlaceholderText = "пусто = как автор";
            b.TextChanged += (_, _) => { if (b.Focused) c.Checked = true; };
            bulk[k] = (c, b);
            t.Controls.Add(c, 0, row);
            t.Controls.Add(b, 1, row);
            row++;
        }
        bulk[Fields.Creator].Chk.Checked = true;
        bulk[Fields.LastModifiedBy].Chk.Checked = true;
        bulk[Fields.Description].Chk.Checked = true;

        chkApp = Opt("app", "Приложение = Microsoft Office (по типу файла)", true,
            "Word → «Microsoft Office Word», PowerPoint → «Microsoft Office PowerPoint», Excel → «Microsoft Excel», версия 16.0000");
        chkPlausible = Opt("plausible", "Правдоподобные редакция и время правки", true,
            "Редакция 3–25, время правки 25–240 минут (но не больше промежутка между датами)");
        AddSpan(t, chkApp, ref row);
        AddSpan(t, chkPlausible, ref row);

        var dates = new Label { Text = "Даты создания и изменения:", AutoSize = true, Margin = new Padding(3, 8, 3, 2) };
        AddSpan(t, dates, ref row);
        rbDateKeep = new RadioButton { Text = "не менять", Checked = true, AutoSize = true };
        rbDateSet = new RadioButton { Text = "задать", AutoSize = true };
        rbDateRand = new RadioButton { Text = "случайно в диапазоне", AutoSize = true };
        AddSpan(t, rbDateKeep, ref row);
        dtCreated = Dtp(DateTime.Now.AddDays(-3).Date.AddHours(10));
        dtModified = Dtp(DateTime.Now.AddDays(-1).Date.AddHours(17).AddMinutes(40));
        AddSpan(t, rbDateSet, ref row);
        AddPair(t, "  создан:", dtCreated, ref row);
        AddPair(t, "  изменён:", dtModified, ref row);
        dtFrom = Dtp(DateTime.Now.AddDays(-14).Date.AddHours(9));
        dtTo = Dtp(DateTime.Now.Date.AddHours(-2));
        AddSpan(t, rbDateRand, ref row);
        AddPair(t, "  с:", dtFrom, ref row);
        AddPair(t, "  по:", dtTo, ref row);
        dtCreated.ValueChanged += (_, _) => rbDateSet.Checked = true;
        dtModified.ValueChanged += (_, _) => rbDateSet.Checked = true;
        dtFrom.ValueChanged += (_, _) => rbDateRand.Checked = true;
        dtTo.ValueChanged += (_, _) => rbDateRand.Checked = true;

        var btns = new FlowLayoutPanel
        {
            AutoSize = true, Margin = new Padding(0, 10, 0, 0), WrapContents = true, MaximumSize = new Size(345, 0),
        };
        var bSel = new Button { Text = "Применить к выбранным", AutoSize = true };
        bSel.Click += (_, _) => ApplyBulk(SelectedEntries());
        var bAll = new Button { Text = "Применить ко всем", AutoSize = true };
        bAll.Click += (_, _) => ApplyBulk(entries);
        var bReset = new Button { Text = "Отменить правки в выбранных", AutoSize = true };
        bReset.Click += (_, _) => ResetEdits(SelectedEntries());
        btns.Controls.AddRange(new Control[] { bSel, bAll, bReset });
        AddSpan(t, btns, ref row);
        side.Controls.Add(Group("Массовая замена", t));

        // --- правка текста ---
        var tx = Table(1);
        var dashRow = new FlowLayoutPanel { AutoSize = true, WrapContents = false, Margin = new Padding(0) };
        var chkDash = Opt("textDash", "Заменить длинные тире «—» на", true,
            "Microsoft Word при наборе « - » сам ставит короткое тире «–». Длинное «—» почти всегда " +
            "вставлено генератором или скопировано.");
        cmbDash = new ComboBox { DropDownStyle = ComboBoxStyle.DropDownList, Width = 110, Margin = new Padding(3, 0, 3, 0) };
        cmbDash.Items.AddRange(new object[] { "«–» тире", "«-» дефис" });
        cmbDash.SelectedIndex = 0;
        dashRow.Controls.AddRange(new Control[] { chkDash, cmbDash });
        tx.Controls.Add(dashRow);
        tx.Controls.Add(Opt("textRename", "    …и в именах файлов", true));
        tx.Controls.Add(Opt("textMd", "Убрать разметку Markdown (**, ##, ```, `код`)", true,
            "Звёздочки и решётки удаляются, сам текст и его форматирование остаются. " +
            "Строки «* пункт» превращаются в «– пункт», строки-разделители ``` и --- удаляются."));
        tx.Controls.Add(Opt("textChat", "Удалить служебные фразы чата", true,
            "Абзацы целиком вида «Конечно! Вот обновлённый вариант…», «Надеюсь, это поможет», " +
            "«Если хотите, я могу…». Каждый удалённый абзац записывается в журнал."));
        tx.Controls.Add(Opt("textPh", "Заполнить [Ваше имя], [Организация], [Дата] из свойств", true,
            "Имя — из поля «Автор», организация — из «Организация», дата — из «Изменён». " +
            "Шаблоны без подходящих данных остаются и перечисляются в журнале."));
        tx.Controls.Add(Opt("altText", "Убрать пути к файлам и подозрительный alt-текст картинок", true,
            "Например «/home/claude/assets/logo.png» — python-pptx записывает путь к картинке в её описание."));
        tx.Controls.Add(new Label
        {
            Text = "Упоминания ИИ в тексте и эмодзи программа только показывает — их правите вы.",
            AutoSize = true, MaximumSize = new Size(340, 0), ForeColor = Color.DimGray, Margin = new Padding(3, 6, 3, 0),
        });
        side.Controls.Add(Group("Правка текста", tx));

        // --- при сохранении ---
        var s = Table(1);
        rbOutFolder = new RadioButton { Text = "Сохранять копии в подпапку «cleaned»", Checked = true, AutoSize = true };
        rbInPlace = new RadioButton { Text = "Перезаписывать исходные файлы", AutoSize = true };
        chkBackup = Opt("backup", "    делать резервную копию .bak", true);
        rbInPlace.CheckedChanged += (_, _) => chkBackup.Enabled = rbInPlace.Checked;
        chkBackup.Enabled = false;
        var modeBox = new FlowLayoutPanel { FlowDirection = FlowDirection.TopDown, AutoSize = true, WrapContents = false };
        modeBox.Controls.AddRange(new Control[] { rbOutFolder, rbInPlace, chkBackup });
        s.Controls.Add(modeBox);

        s.Controls.Add(new Label { Text = "Дополнительная чистка:", AutoSize = true, Margin = new Padding(3, 10, 3, 2) });
        s.Controls.Add(Opt("custom", "Удалить пользовательские свойства", true,
            "docProps/custom.xml — туда генераторы и плагины пишут свои метки"));
        var rev = Opt("reviewers", "Заменить авторов примечаний и исправлений на:", true);
        s.Controls.Add(rev);
        txtReviewer = new TextBox { Width = 320, PlaceholderText = "пусто = автор документа", Margin = new Padding(22, 0, 3, 4) };
        s.Controls.Add(txtReviewer);
        s.Controls.Add(Opt("images", "Очистить метаданные изображений (EXIF, PNG, C2PA)", true,
            "Удаляются EXIF/XMP, текстовые блоки PNG (например, Software: matplotlib) и подписи C2PA, " +
            "которые добавляют ИИ-генераторы картинок. Сама картинка и цветовой профиль не меняются."));
        s.Controls.Add(Opt("embedded", "Чистить свойства встроенных объектов (диаграмм)", true));
        s.Controls.Add(Opt("zip", "Даты внутри архива как у Office (1980-01-01)", true,
            "Microsoft Office записывает в архив нулевые даты, а библиотеки-генераторы — реальное время сборки"));
        s.Controls.Add(Opt("utm", "Убрать метки utm_source=chatgpt.com из ссылок", true));
        s.Controls.Add(Opt("zone", "Снять отметку «Загружено из интернета»", true,
            "Скрытый поток Zone.Identifier: Windows хранит в нём адрес, откуда скачан файл"));
        s.Controls.Add(Opt("filetimes", "Даты файла в Windows = даты документа", true));
        bool officeOk = OfficeSession.AnyAvailable();
        chkOffice = Opt("office", "Пересохранить через Microsoft Office" + (officeOk ? "" : " (не найден)"), false,
            "Каждый файл открывается в Word/PowerPoint/Excel в фоне и сохраняется. Office сам перестраивает " +
            "структуру файла (RSID, settings.xml, тема, миниатюра). Медленнее. Перед запуском закройте эти файлы в Office.");
        chkOffice.Enabled = officeOk;
        s.Controls.Add(chkOffice);

        var bigSave = new Button
        {
            Text = "💾  Сохранить изменения", Height = 38, Width = 330, Margin = new Padding(3, 12, 3, 3),
            Font = new Font(Font.FontFamily, 10f, FontStyle.Bold),
        };
        bigSave.Click += (_, _) => { if (!busy) SaveAll(); };
        s.Controls.Add(bigSave);
        side.Controls.Add(Group("При сохранении", s));
    }

    static void AddSpan(TableLayoutPanel t, Control c, ref int row)
    {
        t.Controls.Add(c, 0, row);
        t.SetColumnSpan(c, 2);
        row++;
    }

    static void AddPair(TableLayoutPanel t, string label, Control c, ref int row)
    {
        t.Controls.Add(new Label { Text = label, AutoSize = true, Margin = new Padding(3, 6, 3, 0) }, 0, row);
        t.Controls.Add(c, 1, row);
        row++;
    }

    // ======================= загрузка файлов =======================

    void OpenFolder()
    {
        using var d = new FolderBrowserDialog { Description = "Папка с документами", UseDescriptionForTitle = true };
        if (d.ShowDialog(this) == DialogResult.OK) AddPaths(new[] { d.SelectedPath });
    }

    void AddFilesDialog()
    {
        using var d = new OpenFileDialog
        {
            Multiselect = true, Title = "Документы Office",
            Filter = "Документы Office|*.docx;*.docm;*.pptx;*.pptm;*.xlsx;*.xlsm|Все файлы|*.*",
        };
        if (d.ShowDialog(this) == DialogResult.OK) AddPaths(d.FileNames);
    }

    public async void AddPaths(IEnumerable<string> paths)
    {
        if (busy) return;
        var list = new List<(string File, string Root)>();
        foreach (var p in paths)
        {
            if (Directory.Exists(p))
                list.AddRange(Scanner.FindFiles(p, chkRecursive.Checked).Select(f => (f, p)));
            else if (File.Exists(p) && Fields.KindFromPath(p) != null)
                list.Add((p, System.IO.Path.GetDirectoryName(System.IO.Path.GetFullPath(p))!));
        }
        var known = entries.Select(e => e.Path).ToHashSet(StringComparer.OrdinalIgnoreCase);
        list = list.Where(x => known.Add(System.IO.Path.GetFullPath(x.File))).ToList();
        if (list.Count == 0) { status.Text = "Новых документов не найдено."; return; }

        var loaded = await RunBusy(list.Count, "Чтение", report =>
            list.Select((x, i) => { var en = LoadEntry(x.File, x.Root); report(i + 1); return en; }).ToList());
        entries.AddRange(loaded);
        RefreshAll();
        Log($"Загружено файлов: {loaded.Count}. С высоким уровнем следов: {loaded.Count(e => e.Result.Score == Severity.High)}.");
    }

    static FileEntry LoadEntry(string path, string root)
    {
        var full = System.IO.Path.GetFullPath(path);
        var r = Inspector.InspectFile(full);
        return new FileEntry
        {
            Path = full, Root = root, Kind = r.Kind, Result = r,
            Original = r.Props.Clone(), Edited = r.Props.Clone(),
        };
    }

    async void Reload()
    {
        if (entries.Count == 0 || !ConfirmDiscard()) return;
        var list = entries.Where(e => File.Exists(e.Path)).Select(e => (e.Path, e.Root)).ToList();
        var loaded = await RunBusy(list.Count, "Чтение", report =>
            list.Select((x, i) => { var en = LoadEntry(x.Path, x.Root); report(i + 1); return en; }).ToList());
        entries.Clear();
        entries.AddRange(loaded);
        RefreshAll();
    }

    bool ConfirmDiscard()
    {
        int dirty = entries.Count(e => e.Dirty);
        if (dirty == 0) return true;
        return MessageBox.Show(this, $"Есть несохранённые правки в {dirty} файл(ах). Отменить их?", "MetaCleaner",
            MessageBoxButtons.YesNo, MessageBoxIcon.Question) == DialogResult.Yes;
    }

    // ======================= таблицы =======================

    void RefreshAll()
    {
        FillProps();
        FillFindings();
        UpdateStatus();
    }

    void FillProps()
    {
        filling = true;
        gridProps.SuspendLayout();
        gridProps.Rows.Clear();
        foreach (var en in entries)
        {
            int i = gridProps.Rows.Add();
            var r = gridProps.Rows[i];
            r.Tag = en;
            r.Cells["file"].Value = System.IO.Path.GetFileName(en.Path);
            r.Cells["file"].ToolTipText = en.Path;
            r.Cells["kind"].Value = Fields.KindName(en.Kind);
            r.Cells["score"].Value = ScoreText(en.Result);
            foreach (var (k, _) in Fields.All)
            {
                r.Cells[k].Value = en.Edited.Get(k);
                if (en.Result.Error != null) r.Cells[k].ReadOnly = true;
            }
            r.Cells["custom"].Value = string.Join(", ", en.Original.Custom.Select(c => c.Name));
            r.Cells["custom"].ToolTipText = string.Join("\n", en.Original.Custom.Select(c => $"{c.Name} = {c.Value}"));
            r.Cells["count"].Value = en.Result.Findings.Count(f => f.Severity >= Severity.Low);
            r.Cells["count"].ToolTipText = string.Join("\n", en.Result.Findings
                .Where(f => f.Severity >= Severity.Medium).Take(12).Select(f => "• " + f.Message));
        }
        gridProps.ResumeLayout();
        filling = false;
    }

    void RefreshPropValues()
    {
        filling = true;
        foreach (DataGridViewRow r in gridProps.Rows)
            if (r.Tag is FileEntry en)
                foreach (var (k, _) in Fields.All) r.Cells[k].Value = en.Edited.Get(k);
        filling = false;
        gridProps.Invalidate();
        UpdateStatus();
    }

    void FillFindings()
    {
        gridFind.Rows.Clear();
        foreach (var en in entries)
            foreach (var f in en.Result.Findings.OrderByDescending(f => f.Severity))
            {
                int i = gridFind.Rows.Add(System.IO.Path.GetFileName(en.Path), Finding.SeverityName(f.Severity),
                    f.Category, f.Message, f.Fix);
                var r = gridFind.Rows[i];
                r.Tag = f;
                r.Cells["file"].ToolTipText = en.Path;
                r.Cells["msg"].ToolTipText = f.Message;
            }
        tabs.TabPages[1].Text = $"Проверка ({entries.Sum(e => e.Result.Findings.Count(f => f.Severity >= Severity.Medium))})";
    }

    void UpdateStatus()
    {
        if (busy) return;
        int hi = entries.Count(e => e.Result.Score == Severity.High);
        int me = entries.Count(e => e.Result.Score == Severity.Medium);
        int dirty = entries.Count(e => e.Dirty);
        status.Text = entries.Count == 0
            ? "Откройте папку или перетащите файлы .docx / .pptx / .xlsx в окно"
            : $"Файлов: {entries.Count}   •   высокий уровень следов: {hi}   •   средний: {me}   •   с несохранёнными правками: {dirty}";
    }

    static string ScoreText(InspectResult r) => r.Error != null ? "● Ошибка" : r.Score switch
    {
        Severity.High => "● Высокий",
        Severity.Medium => "● Средний",
        _ => "● Чисто",
    };

    static Color SevColor(Severity s) => s switch
    {
        Severity.High => ColHigh, Severity.Medium => ColMed, Severity.Low => Color.SteelBlue, _ => ColOk,
    };

    List<FileEntry> SelectedEntries()
    {
        var set = new List<FileEntry>();
        foreach (DataGridViewCell c in gridProps.SelectedCells)
            if (gridProps.Rows[c.RowIndex].Tag is FileEntry en && !set.Contains(en)) set.Add(en);
        foreach (DataGridViewRow r in gridProps.SelectedRows)
            if (r.Tag is FileEntry en && !set.Contains(en)) set.Add(en);
        return set;
    }

    // ======================= массовая замена =======================

    void ApplyBulk(List<FileEntry> targets)
    {
        gridProps.EndEdit();
        targets = targets.Where(e => e.Result.Error == null).ToList();
        if (targets.Count == 0)
        {
            MessageBox.Show(this, "Выберите строки (файлы) в таблице «Свойства».", "MetaCleaner");
            return;
        }
        if (rbDateSet.Checked && dtCreated.Value > dtModified.Value)
        {
            MessageBox.Show(this, "Дата создания позже даты изменения.", "MetaCleaner", MessageBoxButtons.OK, MessageBoxIcon.Warning);
            return;
        }
        if (rbDateRand.Checked && dtTo.Value <= dtFrom.Value.AddHours(2))
        {
            MessageBox.Show(this, "Диапазон дат должен быть больше двух часов.", "MetaCleaner", MessageBoxButtons.OK, MessageBoxIcon.Warning);
            return;
        }

        var rnd = new Random();
        foreach (var en in targets)
        {
            foreach (var k in BulkKeys)
            {
                var (chk, box) = bulk[k];
                if (!chk.Checked) continue;
                var v = box.Text.Trim();
                if (k == Fields.LastModifiedBy && v == "")
                    v = bulk[Fields.Creator].Chk.Checked ? bulk[Fields.Creator].Box.Text.Trim() : en.Edited.Get(Fields.Creator);
                en.Edited[k] = v;
            }

            if (chkApp.Checked)
            {
                en.Edited[Fields.Application] = Fields.ExpectedApplication(en.Kind);
                en.Edited[Fields.AppVersion] = "16.0000";
                if (en.Kind == DocKind.Word && en.Edited.Get(Fields.Template).Trim() == "")
                    en.Edited[Fields.Template] = "Normal.dotm";
            }

            if (rbDateSet.Checked)
            {
                en.Edited[Fields.Created] = dtCreated.Value.ToString(Fields.DateFormat);
                en.Edited[Fields.Modified] = dtModified.Value.ToString(Fields.DateFormat);
            }
            else if (rbDateRand.Checked)
            {
                var from = dtFrom.Value; var to = dtTo.Value;
                double span = (to - from).TotalMinutes;
                var c = from.AddMinutes(rnd.NextDouble() * (span - 120));
                var m = c.AddMinutes(60 + rnd.NextDouble() * ((to - c).TotalMinutes - 60));
                en.Edited[Fields.Created] = c.ToString(Fields.DateFormat);
                en.Edited[Fields.Modified] = m.ToString(Fields.DateFormat);
            }

            if (chkPlausible.Checked)
            {
                if (en.Kind != DocKind.Excel)
                {
                    en.Edited[Fields.Revision] = rnd.Next(3, 26).ToString();
                    int minutes = rnd.Next(25, 241);
                    if (Fields.TryParseLocal(en.Edited.Get(Fields.Created), out var c) &&
                        Fields.TryParseLocal(en.Edited.Get(Fields.Modified), out var m))
                    {
                        int gap = (int)Math.Max(5, (m - c).TotalMinutes);
                        if (minutes > gap) minutes = rnd.Next(Math.Max(1, gap / 3), gap + 1);
                    }
                    en.Edited[Fields.TotalTime] = minutes.ToString();
                }
            }
        }
        RefreshPropValues();
        status.Text = $"Изменения применены к {targets.Count} файл(ам). Проверьте таблицу и нажмите «Сохранить изменения».";
    }

    void ResetEdits(List<FileEntry> targets)
    {
        foreach (var en in targets) en.Edited = en.Original.Clone();
        RefreshPropValues();
    }

    // ======================= профили =======================

    void FillProfiles(string? select = null)
    {
        cmbProfile.Items.Clear();
        cmbProfile.Items.Add("— не выбран —");
        foreach (var p in profiles) cmbProfile.Items.Add(p.Name);
        int idx = select == null ? 0 : profiles.FindIndex(p => p.Name == select) + 1;
        cmbProfile.SelectedIndex = Math.Max(0, idx);
    }

    void SaveProfile()
    {
        var current = cmbProfile.SelectedIndex > 0 ? profiles[cmbProfile.SelectedIndex - 1].Name : "Мой профиль";
        var name = Prompt("Сохранить профиль", "Название профиля:", current);
        if (string.IsNullOrWhiteSpace(name)) return;
        var p = new Profile { Name = name.Trim() };
        foreach (var (k, (chk, box)) in bulk)
        {
            p.Checks["field:" + k] = chk.Checked;
            p.Values[k] = box.Text;
        }
        foreach (var (k, c) in optionChecks) p.Checks[k] = c.Checked;
        p.Values["reviewer"] = txtReviewer.Text;
        p.Values["dash"] = cmbDash.SelectedIndex.ToString();
        p.Checks["inPlace"] = rbInPlace.Checked;
        profiles.RemoveAll(x => x.Name == p.Name);
        profiles.Add(p);
        ProfileStore.Save(profiles);
        FillProfiles(p.Name);
        Log($"Профиль «{p.Name}» сохранён.");
    }

    void DeleteProfile()
    {
        if (cmbProfile.SelectedIndex <= 0) return;
        var p = profiles[cmbProfile.SelectedIndex - 1];
        if (MessageBox.Show(this, $"Удалить профиль «{p.Name}»?", "MetaCleaner", MessageBoxButtons.YesNo) != DialogResult.Yes) return;
        profiles.Remove(p);
        ProfileStore.Save(profiles);
        FillProfiles();
    }

    void ApplyProfile(Profile p)
    {
        foreach (var (k, (chk, box)) in bulk)
        {
            if (p.Checks.TryGetValue("field:" + k, out var on)) chk.Checked = on;
            if (p.Values.TryGetValue(k, out var v)) box.Text = v;
        }
        foreach (var (k, c) in optionChecks)
            if (p.Checks.TryGetValue(k, out var on) && c.Enabled) c.Checked = on;
        if (p.Values.TryGetValue("reviewer", out var r)) txtReviewer.Text = r;
        if (p.Values.TryGetValue("dash", out var ds) && int.TryParse(ds, out var di) && di is 0 or 1) cmbDash.SelectedIndex = di;
        if (p.Checks.TryGetValue("inPlace", out var ip)) { rbInPlace.Checked = ip; rbOutFolder.Checked = !ip; }
        status.Text = $"Загружен профиль «{p.Name}». Выделите файлы и нажмите «Применить».";
    }

    string? Prompt(string title, string label, string value)
    {
        using var f = new Form
        {
            Text = title, Width = 380, Height = 160, FormBorderStyle = FormBorderStyle.FixedDialog,
            StartPosition = FormStartPosition.CenterParent, MaximizeBox = false, MinimizeBox = false, Font = Font,
        };
        var l = new Label { Text = label, Left = 12, Top = 12, AutoSize = true };
        var t = new TextBox { Text = value, Left = 12, Top = 36, Width = 340 };
        var ok = new Button { Text = "OK", DialogResult = DialogResult.OK, Left = 196, Top = 72, Width = 75 };
        var cancel = new Button { Text = "Отмена", DialogResult = DialogResult.Cancel, Left = 277, Top = 72, Width = 75 };
        f.Controls.AddRange(new Control[] { l, t, ok, cancel });
        f.AcceptButton = ok;
        f.CancelButton = cancel;
        return f.ShowDialog(this) == DialogResult.OK ? t.Text : null;
    }

    // ======================= сохранение =======================

    SaveOptions CollectOptions() => new()
    {
        InPlace = rbInPlace.Checked,
        MakeBackup = chkBackup.Checked,
        RemoveCustom = optionChecks["custom"].Checked,
        ReplaceReviewers = optionChecks["reviewers"].Checked,
        Reviewer = txtReviewer.Text.Trim(),
        StripImageMeta = optionChecks["images"].Checked,
        CleanEmbedded = optionChecks["embedded"].Checked,
        OfficeLikeZip = optionChecks["zip"].Checked,
        StripAiUtm = optionChecks["utm"].Checked,
        RemoveZoneId = optionChecks["zone"].Checked,
        SetFileTimes = optionChecks["filetimes"].Checked,
        OfficeResave = chkOffice.Checked && chkOffice.Enabled,
        Text = new TextFixOptions
        {
            FixDashes = optionChecks["textDash"].Checked,
            DashReplacement = cmbDash.SelectedIndex == 1 ? "-" : "–",
            RenameFiles = optionChecks["textRename"].Checked,
            RemoveMarkdown = optionChecks["textMd"].Checked,
            RemoveChatPhrases = optionChecks["textChat"].Checked,
            FillPlaceholders = optionChecks["textPh"].Checked,
            CleanAltText = optionChecks["altText"].Checked,
        },
    };

    void SaveAll()
    {
        gridProps.EndEdit();
        var targets = entries.Where(e => e.Result.Error == null).ToList();
        if (targets.Count == 0) { MessageBox.Show(this, "Нет файлов для сохранения.", "MetaCleaner"); return; }

        // проверка дат и чисел до запуска
        var bad = new List<string>();
        foreach (var en in targets)
            foreach (var k in new[] { Fields.Created, Fields.Modified })
            {
                var v = en.Edited.Get(k);
                if (v != "" && !Fields.TryParseLocal(v, out _))
                    bad.Add($"{System.IO.Path.GetFileName(en.Path)}: «{Fields.Label(k)}» = {v}");
            }
        if (bad.Count > 0)
        {
            MessageBox.Show(this, "Неверный формат дат (нужно ГГГГ-ММ-ДД ЧЧ:ММ):\n\n" + string.Join("\n", bad.Take(10)),
                "MetaCleaner", MessageBoxButtons.OK, MessageBoxIcon.Warning);
            return;
        }

        var o = CollectOptions();
        var where = o.InPlace
            ? (o.MakeBackup ? "Исходные файлы будут перезаписаны (с резервными копиями .bak)." : "Исходные файлы будут перезаписаны БЕЗ резервных копий.")
            : "Очищенные копии будут сохранены в подпапку «cleaned». Исходные файлы не изменятся.";
        var extra = o.OfficeResave ? "\n\nФайлы будут пересохранены через Microsoft Office — закройте их в Office перед продолжением." : "";
        if (MessageBox.Show(this, $"Сохранить {targets.Count} файл(ов)?\n\n{where}{extra}", "MetaCleaner",
                MessageBoxButtons.OKCancel, MessageBoxIcon.Question) != DialogResult.OK) return;

        var jobs = targets.Select(e => (Entry: e, Props: e.Edited.Clone())).ToList();
        var results = new List<(FileEntry Entry, SaveResult? Res, string? Error, InspectResult? After)>();
        SetBusy(true, jobs.Count, "Сохранение");
        Log($"--- Сохранение {jobs.Count} файл(ов) ---");

        var thread = new Thread(() =>
        {
            OfficeSession? office = o.OfficeResave ? new OfficeSession() : null;
            try
            {
                int i = 0;
                foreach (var (en, props) in jobs)
                {
                    try
                    {
                        var r = Cleaner.SaveFile(en.Path, en.Root, props, o, office);
                        results.Add((en, r, null, Inspector.InspectFile(r.OutputPath)));
                    }
                    catch (Exception ex) { results.Add((en, null, ex.Message, null)); }
                    int done = ++i;
                    BeginInvoke((MethodInvoker)(() => progress.Value = Math.Min(progress.Maximum, done)));
                }
            }
            finally { office?.Dispose(); }
            BeginInvoke((MethodInvoker)(() => SaveDone(results, o)));
        });
        thread.SetApartmentState(ApartmentState.STA); // нужно для автоматизации Office
        thread.IsBackground = true;
        thread.Start();
    }

    void SaveDone(List<(FileEntry Entry, SaveResult? Res, string? Error, InspectResult? After)> results, SaveOptions o)
    {
        SetBusy(false, 0, "");
        int ok = 0, fail = 0;
        foreach (var (en, res, err, after) in results)
        {
            var name = System.IO.Path.GetFileName(en.Path);
            if (err != null || res == null)
            {
                fail++;
                Log($"✖ {name}: ОШИБКА — {err}. Исходный файл не изменён.");
                continue;
            }
            ok++;
            Log($"✔ {name} → {res.OutputPath}");
            foreach (var a in res.Actions) Log("    • " + a);
            if (res.BackupPath != null) Log("    • резервная копия: " + res.BackupPath);
            if (after != null)
            {
                var left = after.Findings.Where(f => f.Severity >= Severity.Medium).ToList();
                Log(left.Count == 0
                    ? "    Проверка после сохранения: значимых следов не найдено."
                    : $"    Проверка после сохранения — осталось {left.Count}:");
                foreach (var f in left) Log($"      [{Finding.SeverityName(f.Severity)}] {f.Message} — {f.Fix}");
                if (o.InPlace)
                {
                    en.Path = res.OutputPath;
                    en.Result = after;
                    en.Original = after.Props.Clone();
                    en.Edited = after.Props.Clone();
                }
            }
        }
        Log($"--- Готово: {ok} успешно, {fail} с ошибками ---");
        tabs.SelectedIndex = 2;

        if (o.InPlace) { RefreshAll(); return; }
        RefreshAll();
        if (ok == 0) return;
        var outDirs = results.Where(r => r.Res != null).Select(r => System.IO.Path.Combine(r.Entry.Root, o.OutputFolderName))
            .Distinct(StringComparer.OrdinalIgnoreCase).ToList();
        var ans = MessageBox.Show(this,
            $"Готово: {ok} файл(ов) сохранено в «{o.OutputFolderName}»" + (fail > 0 ? $", ошибок: {fail}" : "") +
            ".\n\nДа — загрузить очищенные копии в таблицу для проверки.\nНет — открыть папку в проводнике.",
            "MetaCleaner", MessageBoxButtons.YesNoCancel, MessageBoxIcon.Information);
        if (ans == DialogResult.Yes)
        {
            foreach (var en in entries) en.Edited = en.Original.Clone();
            entries.Clear();
            RefreshAll();
            AddPaths(outDirs);
        }
        else if (ans == DialogResult.No)
        {
            foreach (var d in outDirs.Take(3))
                try { Process.Start(new ProcessStartInfo("explorer.exe", $"\"{d}\"") { UseShellExecute = true }); } catch { }
        }
    }

    // ======================= служебное =======================

    async Task<T> RunBusy<T>(int total, string what, Func<Action<int>, T> work)
    {
        SetBusy(true, total, what);
        try
        {
            return await Task.Run(() => work(n => BeginInvoke((MethodInvoker)(() =>
            {
                progress.Value = Math.Min(progress.Maximum, n);
                status.Text = $"{what}: {n} из {total}…";
            }))));
        }
        finally { SetBusy(false, 0, ""); }
    }

    void SetBusy(bool on, int total, string what)
    {
        busy = on;
        progress.Visible = on;
        if (on)
        {
            progress.Maximum = Math.Max(1, total);
            progress.Value = 0;
            status.Text = what + "…";
            Cursor = Cursors.WaitCursor;
        }
        else
        {
            Cursor = Cursors.Default;
            UpdateStatus();
        }
        toolbar.Enabled = !on;
        side.Enabled = !on;
        gridProps.ReadOnly = on;
    }

    void Log(string line) => logBox.AppendText($"[{DateTime.Now:HH:mm:ss}] {line}{Environment.NewLine}");

    void ExportCsv()
    {
        var g = tabs.SelectedIndex == 1 ? gridFind : gridProps;
        if (g.Rows.Count == 0) { MessageBox.Show(this, "Таблица пуста.", "MetaCleaner"); return; }
        using var d = new SaveFileDialog
        {
            Filter = "CSV (Excel)|*.csv",
            FileName = (tabs.SelectedIndex == 1 ? "проверка" : "свойства") + $"_{DateTime.Now:yyyy-MM-dd_HHmm}.csv",
        };
        if (d.ShowDialog(this) != DialogResult.OK) return;
        static string Q(object? v) => "\"" + (v?.ToString() ?? "").Replace("\"", "\"\"") + "\"";
        var sb = new StringBuilder();
        var cols = g.Columns.Cast<DataGridViewColumn>().OrderBy(c => c.DisplayIndex).ToList();
        if (g == gridProps) sb.AppendLine(string.Join(";", new[] { Q("Путь") }.Concat(cols.Select(c => Q(c.HeaderText)))));
        else sb.AppendLine(string.Join(";", cols.Select(c => Q(c.HeaderText))));
        foreach (DataGridViewRow r in g.Rows)
        {
            var cells = cols.Select(c => Q(r.Cells[c.Index].Value));
            if (g == gridProps && r.Tag is FileEntry en) cells = new[] { Q(en.Path) }.Concat(cells);
            sb.AppendLine(string.Join(";", cells));
        }
        File.WriteAllText(d.FileName, sb.ToString(), new UTF8Encoding(true));
        Log("Экспортировано: " + d.FileName);
    }
}
