# Библиотекарь — каталог и ручная сортировка файлов для Windows PowerShell 5.1+
# Версия 16: план раскладки можно выполнить (только строки с «Перенести = да»), с журналом и отменой последнего переноса.
# Папки переезжают целиком, сохраняя структуру; по типам раскладываются только файлы без папки.
# Каталог и поиск. Точные дубликаты отправляются в Корзину только после общего подтверждения.
# Еженедельная уборка «Загрузок» и Рабочего стола: кнопка «Еженедельная уборка» или запуск с ключом -WeeklyCleanup
# (его вызывает Планировщик заданий). Уверенное — сама, спорное — с галочками, свежее и личное не трогает.
param([switch]$WeeklyCleanup, [switch]$Demo)

Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing
Add-Type -AssemblyName Microsoft.VisualBasic
Add-Type -AssemblyName System.IO.Compression.FileSystem
[System.Windows.Forms.Application]::EnableVisualStyles()

# Own projects and rules live in «настройки.json» next to the program, not in the code:
# the same program works for anyone, each person keeps their own settings file.
$script:SettingsPath = Join-Path $PSScriptRoot 'настройки.json'
$script:Projects = @(); $script:ExtraRules = @(); $script:PromptLibDir = ''; $dataFolderName = 'LibrarianFiles'
if (Test-Path -LiteralPath $script:SettingsPath) {
    try {
        $cfg = [IO.File]::ReadAllText($script:SettingsPath, [Text.Encoding]::UTF8) | ConvertFrom-Json
        $script:Projects = @($cfg.'Проекты' | ForEach-Object { $_ } | Where-Object { $_.'Слова' -and $_.'Папка' })
        $script:ExtraRules = @($cfg.'Свои правила' | ForEach-Object { $_ } | Where-Object { $_.'Слова' -and $_.'Полка' })
        $lib = [string]$cfg.'Библиотека промптов'
        if ($lib) { $script:PromptLibDir = if ([IO.Path]::IsPathRooted($lib)) { $lib } else { [IO.Path]::GetFullPath((Join-Path $PSScriptRoot $lib)) } }
        if ($cfg.'Папка данных') { $dataFolderName = [string]$cfg.'Папка данных' }
    } catch { }
}
# The prompt library (Telegram channel + site) is shown only where it exists.
$script:HasPrompts = [bool]($script:PromptLibDir -and (Test-Path -LiteralPath $script:PromptLibDir -PathType Container))
$script:DataHome = Join-Path $env:LOCALAPPDATA $dataFolderName
$script:ConfigPath = Join-Path $script:DataHome 'roots.json'
$script:RootsMigrationPath = Join-Path $script:DataHome 'roots_v6_migrated.txt'
$script:IndexPath = Join-Path $script:DataHome 'index.csv'
$script:DuplicateGroups = @()
$script:HashScanComplete = $false
$script:DuplicateReportPath = Join-Path $script:DataHome 'duplicates_report.txt'
$script:CleanupLogPath = Join-Path $script:DataHome 'duplicate_cleanup_log.txt'
$script:ContentReadErrors = 0
$script:LastContentReadError = ''
$script:ContentReadReportPath = Join-Path $script:DataHome 'content_read_report.txt'
$script:OrganizationPlanPath = Join-Path $script:DataHome 'organization_plan.html'
$script:OrganizationPlanCsvPath = Join-Path $script:DataHome 'organization_plan.csv'
$script:StorageRootConfigPath = Join-Path $script:DataHome 'storage_root.txt'
# Root folder for the sorted groups (01_РАБОТА, 02_МОИ_ПРОЕКТЫ…). Old plans used МОЁ_ХРАНИЛИЩЕ.
$script:LibraryRootName = 'КНИЖНЫЕ_ПОЛКИ'
$script:OldLibraryRootNames = @('МОЁ_ХРАНИЛИЩЕ')
$script:MoveLogDir = Join-Path $script:DataHome 'move_logs'
$script:WordApp = $null
$script:WordMacrosDisabled = $false
$script:WeeklyMode = [bool]$WeeklyCleanup
$script:CleanupAgeDays = 7      # files younger than this are probably still in use
$script:InstallerAgeDays = 14   # older installers go to the Recycle Bin
$script:MediaDir = Join-Path $PSScriptRoot 'media'
$script:SweepImage = Join-Path $script:MediaDir 'uborka.png'
$script:SweepVideo = Join-Path $script:MediaDir 'uborka.mp4'
$script:WeeklyReportDir = Join-Path $script:DataHome 'weekly_cleanup'
New-Item -ItemType Directory -Path $script:DataHome -Force | Out-Null

function Get-DefaultRoots {
    $roots = New-Object System.Collections.Generic.List[string]
    $specialFolders = @('Desktop','MyDocuments','MyPictures','MyVideos','MyMusic')
    foreach ($name in $specialFolders) {
        try {
            $special = [Enum]::Parse([System.Environment+SpecialFolder], $name)
            $path = [Environment]::GetFolderPath($special)
            if ($path -and (Test-Path -LiteralPath $path -PathType Container)) { $roots.Add($path) }
        } catch { }
    }
    $downloads = $null
    try {
        $userShell = Get-ItemProperty -LiteralPath 'HKCU:\Software\Microsoft\Windows\CurrentVersion\Explorer\User Shell Folders' -ErrorAction Stop
        $downloadGuid = '{374DE290-123F-4565-9164-39C4925E467B}'
        $downloads = [Environment]::ExpandEnvironmentVariables([string]$userShell.$downloadGuid)
    } catch { }
    if (!$downloads) { $downloads = Join-Path $env:USERPROFILE 'Downloads' }
    if (Test-Path -LiteralPath $downloads -PathType Container) { $roots.Add($downloads) }
    return @($roots | Select-Object -Unique)
}
function Get-Roots {
    $defaults = @(Get-DefaultRoots)
    if (Test-Path -LiteralPath $script:ConfigPath) {
        try {
            # ConvertFrom-Json in PowerShell 5.1 emits a JSON array as one object: enumerate it explicitly,
            # otherwise all folders were glued into one string on the next save.
            $saved = @((Get-Content -LiteralPath $script:ConfigPath -Raw -Encoding UTF8 | ConvertFrom-Json) | ForEach-Object { $_ } |
                Where-Object { $_ -is [string] -and (Test-Path -LiteralPath $_ -PathType Container) } | Select-Object -Unique)
            # One-time migration: add the six folders selected for the first full inventory,
            # while retaining any folders the user had already added.
            if (!(Test-Path -LiteralPath $script:RootsMigrationPath)) {
                $saved = @(@($saved) + @($defaults) | Select-Object -Unique)
                Save-Roots $saved
                Set-Content -LiteralPath $script:RootsMigrationPath -Value 'Completed' -Encoding UTF8
            }
            return @($saved)
        } catch { }
    }
    return $defaults
}
function Save-Roots($Roots) {
    $json = ConvertTo-Json -InputObject @($Roots) -Depth 3
    Set-Content -LiteralPath $script:ConfigPath -Value $json -Encoding UTF8
}
function Format-Bytes([long]$Bytes) {
    if ($Bytes -ge 1GB) { return ('{0:N2} ГБ' -f ($Bytes / 1GB)) }
    if ($Bytes -ge 1MB) { return ('{0:N1} МБ' -f ($Bytes / 1MB)) }
    if ($Bytes -ge 1KB) { return ('{0:N0} КБ' -f ($Bytes / 1KB)) }
    return ("$Bytes байт")
}
function Get-Category([string]$Extension) {
    $e = $Extension.ToLowerInvariant()
    if ($e -in @('.jpg','.jpeg','.png','.gif','.webp','.bmp','.tif','.tiff','.heic','.svg')) { return 'Изображения' }
    if ($e -in @('.mp4','.mov','.avi','.mkv','.wmv','.webm')) { return 'Видео' }
    if ($e -in @('.mp3','.wav','.m4a','.aac','.flac','.ogg')) { return 'Аудио' }
    if ($e -in @('.doc','.docx','.rtf','.odt','.txt','.md')) { return 'Тексты' }
    if ($e -in @('.xls','.xlsx','.csv','.ods')) { return 'Таблицы' }
    if ($e -in @('.ppt','.pptx','.odp')) { return 'Презентации' }
    if ($e -in @('.pdf','.djvu')) { return 'PDF и книги' }
    if ($e -in @('.zip','.7z','.rar','.tar','.gz')) { return 'Архивы' }
    if ($e -in @('.psd','.ai','.aep','.blend','.fig','.indd')) { return 'Дизайн-проекты' }
    if ($e -in @('.py','.ps1','.js','.html','.css','.json','.yml','.yaml')) { return 'Код и данные' }
    return 'Другие файлы'
}

$form = New-Object System.Windows.Forms.Form
$form.Text = 'Моя файловая библиотека'
$form.StartPosition = 'CenterScreen'
$form.Size = [System.Drawing.Size]::new(1230,760)
$form.MinimumSize = [System.Drawing.Size]::new(980,590)
$form.Font = [System.Drawing.Font]::new('Segoe UI', 9)
$iconPath = Join-Path $PSScriptRoot 'media\librarian.ico'
if (Test-Path -LiteralPath $iconPath) { try { $form.Icon = New-Object System.Drawing.Icon $iconPath } catch { } }

$top = New-Object System.Windows.Forms.Panel
$top.Height = 150

$rootsLabel = New-Object System.Windows.Forms.Label
$rootsLabel.Text = 'Папки для каталога:'; $rootsLabel.Location = [System.Drawing.Point]::new(12,10); $rootsLabel.AutoSize = $true
$top.Controls.Add($rootsLabel)
$rootsBox = New-Object System.Windows.Forms.ListBox
$rootsBox.Location = [System.Drawing.Point]::new(12,32); $rootsBox.Size = [System.Drawing.Size]::new(580,58)
$rootsBox.HorizontalScrollbar = $true
$rootsBox.TabIndex = 10
$top.Controls.Add($rootsBox)

$addRoot = New-Object System.Windows.Forms.Button
$addRoot.Text = 'Добавить папку…'; $addRoot.Location = [System.Drawing.Point]::new(602,32); $addRoot.Size = [System.Drawing.Size]::new(130,28)
$addRoot.TabIndex = 11
$top.Controls.Add($addRoot)
$removeRoot = New-Object System.Windows.Forms.Button
$removeRoot.Text = 'Убрать из списка'; $removeRoot.Location = [System.Drawing.Point]::new(602,63); $removeRoot.Size = [System.Drawing.Size]::new(130,28)
$removeRoot.TabIndex = 12
$top.Controls.Add($removeRoot)
$scanButton = New-Object System.Windows.Forms.Button
$scanButton.Text = 'Обновить каталог'; $scanButton.Location = [System.Drawing.Point]::new(740,32); $scanButton.Size = [System.Drawing.Size]::new(145,28)
$scanButton.TabIndex = 13
$top.Controls.Add($scanButton)
$dupeButton = New-Object System.Windows.Forms.Button
$dupeButton.Text = 'Найти дубликаты'; $dupeButton.Location = [System.Drawing.Point]::new(740,63); $dupeButton.Size = [System.Drawing.Size]::new(145,28)
$dupeButton.TabIndex = 14
$top.Controls.Add($dupeButton)
$reportButton = New-Object System.Windows.Forms.Button
$reportButton.Text = 'Отчёт о дубликатах'; $reportButton.Location = [System.Drawing.Point]::new(892,63); $reportButton.Size = [System.Drawing.Size]::new(145,28)
$reportButton.TabIndex = 15
$top.Controls.Add($reportButton)
$contentButton = New-Object System.Windows.Forms.Button
$contentButton.Text = 'Читать содержание'; $contentButton.Location = [System.Drawing.Point]::new(892,32); $contentButton.Size = [System.Drawing.Size]::new(145,28)
$contentButton.TabIndex = 16
$top.Controls.Add($contentButton)
$weeklyButton = New-Object System.Windows.Forms.Button
$weeklyButton.Text = 'Еженедельная уборка'; $weeklyButton.Location = [System.Drawing.Point]::new(1045,32); $weeklyButton.Size = [System.Drawing.Size]::new(150,59)
$weeklyButton.BackColor = [System.Drawing.Color]::FromArgb(255,241,214)
$weeklyButton.TabIndex = 17

$searchLabel = New-Object System.Windows.Forms.Label
$searchLabel.Text = 'Что найти? Напишите или нажмите «Сказать» — например: «договор с Ивановым», «фото собаки», «презентация про отпуск»:'; $searchLabel.Location = [System.Drawing.Point]::new(12,96); $searchLabel.AutoSize = $true
$top.Controls.Add($searchLabel)
$searchBox = New-Object System.Windows.Forms.TextBox
$searchBox.Location = [System.Drawing.Point]::new(12,116); $searchBox.Size = [System.Drawing.Size]::new(440,27)
$searchBox.TabIndex = 0
$searchBox.Font = [System.Drawing.Font]::new('Segoe UI', 10)
$top.Controls.Add($searchBox)
$searchTip = New-Object System.Windows.Forms.ToolTip
$searchTip.SetToolTip($searchBox, 'Введите запрос и нажмите Enter или кнопку «Найти». Поле поиска получает фокус при запуске.')
$searchButton = New-Object System.Windows.Forms.Button
$searchButton.Text = 'Найти'; $searchButton.Location = [System.Drawing.Point]::new(640,115); $searchButton.Size = [System.Drawing.Size]::new(80,28)
$searchButton.TabIndex = 1
$top.Controls.Add($searchButton)
$resetButton = New-Object System.Windows.Forms.Button
$resetButton.Text = 'Сбросить'; $resetButton.Location = [System.Drawing.Point]::new(726,115); $resetButton.Size = [System.Drawing.Size]::new(80,28)
$resetButton.TabIndex = 2
$top.Controls.Add($resetButton)
$form.AcceptButton = $searchButton
$excelPlanButton = New-Object System.Windows.Forms.Button
$excelPlanButton.Text = 'Открыть план в Excel'; $excelPlanButton.Location = [System.Drawing.Point]::new(812,115); $excelPlanButton.Size = [System.Drawing.Size]::new(140,28)
$excelPlanButton.TabIndex = 3
$top.Controls.Add($excelPlanButton)
$searchTip.SetToolTip($excelPlanButton, 'Открывает таблицу плана в Excel. Там можно поставить «да»/«нет», сохранить, закрыть Excel и нажать «Выполнить план…».')
$micButton = New-Object System.Windows.Forms.Button
$micButton.Text = 'Сказать'; $micButton.Location = [System.Drawing.Point]::new(960,115); $micButton.Size = [System.Drawing.Size]::new(105,28)
$micButton.BackColor = [System.Drawing.Color]::FromArgb(255,241,214)
$top.Controls.Add($micButton)
$searchTip.SetToolTip($micButton, 'Включает голосовой ввод Windows (то же, что Win+H). Скажите, что найти, — поиск начнётся сам, ответ появится текстом над списком.')
$voiceCheck = New-Object System.Windows.Forms.CheckBox
$voiceCheck.Text = 'Голос Артура'; $voiceCheck.Location = [System.Drawing.Point]::new(1072,119); $voiceCheck.AutoSize = $true; $voiceCheck.Checked = $true
$top.Controls.Add($voiceCheck)
# Two-way switch: «Поиск» only searches, «Задание» carries out instructions (delete, open, cleanup).
$modeFind = New-Object System.Windows.Forms.RadioButton
$modeFind.Appearance = 'Button'; $modeFind.TextAlign = 'MiddleCenter'; $modeFind.Text = 'Поиск'; $modeFind.Checked = $true
$modeFind.Location = [System.Drawing.Point]::new(460,115); $modeFind.Size = [System.Drawing.Size]::new(80,28)
$modeTask = New-Object System.Windows.Forms.RadioButton
$modeTask.Appearance = 'Button'; $modeTask.TextAlign = 'MiddleCenter'; $modeTask.Text = 'Задание'
$modeTask.Location = [System.Drawing.Point]::new(542,115); $modeTask.Size = [System.Drawing.Size]::new(90,28)
$top.Controls.Add($modeFind); $top.Controls.Add($modeTask)
$searchTip.SetToolTip($modeFind, 'Режим поиска: Библиотекарь только ищет и ничего не удаляет.')
$searchTip.SetToolTip($modeTask, 'Режим задания: удалить или открыть папку, сделать уборку. Перед удалением он всегда спросит.')

$grid = New-Object System.Windows.Forms.ListView
$grid.Dock = 'Fill'; $grid.View = [System.Windows.Forms.View]::Details
$grid.FullRowSelect = $true; $grid.MultiSelect = $true; $grid.GridLines = $true
$grid.HideSelection = $false; $grid.UseCompatibleStateImageBehavior = $false
$grid.BackColor = [System.Drawing.Color]::White
$grid.Columns.Add('Имя файла',245) | Out-Null
$grid.Columns.Add('Категория',115) | Out-Null
$grid.Columns.Add('Тип',65) | Out-Null
$grid.Columns.Add('Размер, МБ',85) | Out-Null
$grid.Columns.Add('Размер, байт',105) | Out-Null
$grid.Columns.Add('Создан',130) | Out-Null
$grid.Columns.Add('Изменён',130) | Out-Null
$grid.Columns.Add('Фрагмент содержания',300) | Out-Null
$grid.Columns.Add('Результат чтения',150) | Out-Null
$pathPanel = New-Object System.Windows.Forms.Panel
$pathPanel.Height = 36
# Anchors are measured from the panel's initial width, so start it at the real width.
$pathPanel.Width = $form.ClientSize.Width
$pathLabel = New-Object System.Windows.Forms.Label
$pathLabel.Text = 'Полный путь выбранного файла:'; $pathLabel.Location = [System.Drawing.Point]::new(12,9); $pathLabel.AutoSize = $true
$pathPanel.Controls.Add($pathLabel)
$selectedPathBox = New-Object System.Windows.Forms.TextBox
$selectedPathBox.ReadOnly = $true; $selectedPathBox.Location = [System.Drawing.Point]::new(215,5); $selectedPathBox.Size = [System.Drawing.Size]::new(($pathPanel.Width - 215 - 125),25)
$selectedPathBox.Anchor = 'Top,Left,Right'; $selectedPathBox.BackColor = [System.Drawing.Color]::White
$pathPanel.Controls.Add($selectedPathBox)
$copyPathButton = New-Object System.Windows.Forms.Button
$copyPathButton.Text = 'Копировать'; $copyPathButton.Location = [System.Drawing.Point]::new(($pathPanel.Width - 117),4); $copyPathButton.Size = [System.Drawing.Size]::new(105,27); $copyPathButton.Anchor = 'Top,Right'
$pathPanel.Controls.Add($copyPathButton)

$bottom = New-Object System.Windows.Forms.Panel
$bottom.Height = 48
$openButton = New-Object System.Windows.Forms.Button
$openButton.Text = 'Открыть файл'; $openButton.Location = [System.Drawing.Point]::new(12,9); $openButton.Size = [System.Drawing.Size]::new(115,29)
$bottom.Controls.Add($openButton)
$folderButton = New-Object System.Windows.Forms.Button
$folderButton.Text = 'Показать в папке'; $folderButton.Location = [System.Drawing.Point]::new(134,9); $folderButton.Size = [System.Drawing.Size]::new(130,29)
$bottom.Controls.Add($folderButton)
$organizeButton = New-Object System.Windows.Forms.Button
$organizeButton.Text = 'Выполнить план…'; $organizeButton.Location = [System.Drawing.Point]::new(408,9); $organizeButton.Size = [System.Drawing.Size]::new(140,29)
$bottom.Controls.Add($organizeButton)
$planButton = New-Object System.Windows.Forms.Button
$planButton.Text = 'План раскладки'; $planButton.Location = [System.Drawing.Point]::new(271,9); $planButton.Size = [System.Drawing.Size]::new(130,29)
$bottom.Controls.Add($planButton)
$undoMoveButton = New-Object System.Windows.Forms.Button
$undoMoveButton.Text = 'Отменить перенос'; $undoMoveButton.Location = [System.Drawing.Point]::new(555,9); $undoMoveButton.Size = [System.Drawing.Size]::new(150,29)
$bottom.Controls.Add($undoMoveButton)
$bottomTip = New-Object System.Windows.Forms.ToolTip
$bottomTip.SetToolTip($planButton, 'Если в таблице выбраны строки — план только для них; иначе — для всех файлов Рабочего стола. Ничего не перемещает.')
$bottomTip.SetToolTip($organizeButton, 'Переносит только строки плана, где в колонке «Перенести» стоит «да». Перед переносом показывает итог и спрашивает подтверждение.')
$bottomTip.SetToolTip($undoMoveButton, 'Возвращает файлы последнего переноса на прежние места по журналу.')
$cleanupButton = New-Object System.Windows.Forms.Button
$cleanupButton.Text = 'Убрать точные дубли'; $cleanupButton.Location = [System.Drawing.Point]::new(712,9); $cleanupButton.Size = [System.Drawing.Size]::new(170,29)
$cleanupButton.Enabled = $false
$bottom.Controls.Add($cleanupButton)
$status = New-Object System.Windows.Forms.Label
$status.Text = 'Каталог читает сведения о файлах. Ничего не перемещается автоматически.'
$status.AutoEllipsis = $true; $status.Location = [System.Drawing.Point]::new(890,14); $status.Size = [System.Drawing.Size]::new(440,22); $status.Anchor = 'Top,Left,Right'
$bottom.Controls.Add($status)

# Separate rows keep the search, results, selected path, and actions from overlapping.
$layout = New-Object System.Windows.Forms.TableLayoutPanel
$layout.Dock = 'Fill'; $layout.ColumnCount = 1; $layout.RowCount = 5
$layout.Margin = [System.Windows.Forms.Padding]::Empty
$layout.Padding = [System.Windows.Forms.Padding]::Empty
$layout.ColumnStyles.Add([System.Windows.Forms.ColumnStyle]::new([System.Windows.Forms.SizeType]::Percent,100)) | Out-Null
$layout.RowStyles.Add([System.Windows.Forms.RowStyle]::new([System.Windows.Forms.SizeType]::Absolute,150)) | Out-Null
$layout.RowStyles.Add([System.Windows.Forms.RowStyle]::new([System.Windows.Forms.SizeType]::Absolute,154)) | Out-Null
$layout.RowStyles.Add([System.Windows.Forms.RowStyle]::new([System.Windows.Forms.SizeType]::Percent,100)) | Out-Null
$layout.RowStyles.Add([System.Windows.Forms.RowStyle]::new([System.Windows.Forms.SizeType]::Absolute,36)) | Out-Null
$layout.RowStyles.Add([System.Windows.Forms.RowStyle]::new([System.Windows.Forms.SizeType]::Absolute,48)) | Out-Null
# The Librarian's written answer: what was found and where it lies.
$answerLabel = New-Object System.Windows.Forms.Label
$answerLabel.Font = [System.Drawing.Font]::new('Segoe UI', 12)
$answerLabel.ForeColor = [System.Drawing.Color]::FromArgb(76,59,37)
$answerLabel.BackColor = [System.Drawing.Color]::FromArgb(255,248,232)
$answerLabel.Padding = [System.Windows.Forms.Padding]::new(12,4,12,4)
$answerLabel.AutoEllipsis = $true
$answerLabel.Text = 'Напишите, что найти, или нажмите «Сказать». Я отвечу здесь.'
# The Librarian himself: short animated scenes from the cartoon (media\helper\*.gif) with Arthur's lines.
$helperBox = New-Object System.Windows.Forms.PictureBox
$helperBox.Dock = 'Left'; $helperBox.Width = 272; $helperBox.SizeMode = 'Zoom'
$helperBox.BackColor = [System.Drawing.Color]::FromArgb(255,248,232)
$answerPanel = New-Object System.Windows.Forms.Panel
$answerLabel.Dock = 'Fill'
$answerPanel.Controls.Add($answerLabel)
$answerPanel.Controls.Add($helperBox)
# Big jobs live next to the Librarian: weekly cleanup, prompts from the Telegram channel, the prompt site.
$jobsPanel = New-Object System.Windows.Forms.FlowLayoutPanel
$jobsPanel.Dock = 'Right'; $jobsPanel.Width = 186; $jobsPanel.FlowDirection = 'TopDown'; $jobsPanel.WrapContents = $false
$jobsPanel.Padding = [System.Windows.Forms.Padding]::new(6,6,6,4); $jobsPanel.BackColor = [System.Drawing.Color]::FromArgb(255,248,232)
$weeklyButton.Size = [System.Drawing.Size]::new(170,42); $weeklyButton.Margin = [System.Windows.Forms.Padding]::new(0,0,0,4)
$promptsButton = New-Object System.Windows.Forms.Button
$promptsButton.Text = 'Промпты из канала'; $promptsButton.Size = [System.Drawing.Size]::new(170,42); $promptsButton.Margin = [System.Windows.Forms.Padding]::new(0,0,0,4)
$promptsButton.BackColor = [System.Drawing.Color]::FromArgb(255,241,214)
$siteButton = New-Object System.Windows.Forms.Button
$siteButton.Text = 'Сайт промптов'; $siteButton.Size = [System.Drawing.Size]::new(170,42); $siteButton.Margin = [System.Windows.Forms.Padding]::new(0)
$siteButton.BackColor = [System.Drawing.Color]::FromArgb(255,241,214)
$jobsPanel.Controls.AddRange(@($weeklyButton, $promptsButton, $siteButton))
if (!$script:HasPrompts) { $promptsButton.Visible = $false; $siteButton.Visible = $false }
function Get-Plaques {
    # Wooden plaques on the right; the prompt ones only where the prompt library exists.
    if ($script:HasPrompts) { return @($weeklyButton, $promptsButton, $siteButton, $stageMore) }
    return @($weeklyButton, $stageMore)
}
$answerPanel.Controls.Add($jobsPanel)
$top.Dock = 'Fill'; $answerPanel.Dock = 'Fill'; $grid.Dock = 'Fill'; $pathPanel.Dock = 'Fill'; $bottom.Dock = 'Fill'
$layout.Controls.Add($top,0,0)
$layout.Controls.Add($answerPanel,0,1)
$layout.Controls.Add($grid,0,2)
$layout.Controls.Add($pathPanel,0,3)
$layout.Controls.Add($bottom,0,4)
$form.Controls.Add($layout)

# ================= Оформление «старая библиотека»: шапка-картина, пергамент, дерево и золото =================
# Rare buttons leave the screen (they stay alive for F-keys and the «Ещё» menu); the logic is untouched.
$script:Ink = [System.Drawing.Color]::FromArgb(64,44,26)          # dark brown text
$script:Parchment = [System.Drawing.Color]::FromArgb(247,239,222)  # page
$script:Paper = [System.Drawing.Color]::FromArgb(253,249,239)      # lighter page
$script:Wood = [System.Drawing.Color]::FromArgb(110,70,38)         # wooden buttons
$script:Gold = [System.Drawing.Color]::FromArgb(196,150,72)        # frames
$script:Cream = [System.Drawing.Color]::FromArgb(255,241,212)      # light text on wood
$script:TitleFont = [System.Drawing.Font]::new('Georgia', 22, [System.Drawing.FontStyle]::Bold)
$script:BookFont = [System.Drawing.Font]::new('Georgia', 11.5)
function Set-WoodButton($Button, [switch]$Light) {
    $Button.FlatStyle = 'Flat'; $Button.FlatAppearance.BorderColor = $script:Gold; $Button.FlatAppearance.BorderSize = 1
    $Button.Cursor = [System.Windows.Forms.Cursors]::Hand
    if ($Light) { $Button.BackColor = $script:Paper; $Button.ForeColor = $script:Ink; $Button.FlatAppearance.MouseOverBackColor = [System.Drawing.Color]::FromArgb(250,232,196) }
    else { $Button.BackColor = $script:Wood; $Button.ForeColor = $script:Cream; $Button.FlatAppearance.MouseOverBackColor = [System.Drawing.Color]::FromArgb(140,92,50) }
    if ($Button -is [System.Windows.Forms.RadioButton]) { $Button.FlatAppearance.CheckedBackColor = $script:Gold }
}
# ---- Messages in the library style instead of the grey Windows boxes ----
# Every [LibrarianApp.LibBox]::Show(...) comes here (same arguments as MessageBox.Show).
if (-not ('LibrarianApp.LibBox' -as [type])) {
    Add-Type -ReferencedAssemblies System.Windows.Forms -TypeDefinition @'
namespace LibrarianApp {
    using System.Windows.Forms;
    public static class LibBox {
        public static System.Func<string, string, string, string, string> Impl;
        public static DialogResult Show(string text) { return Show(text, "", MessageBoxButtons.OK, MessageBoxIcon.None); }
        public static DialogResult Show(string text, string caption) { return Show(text, caption, MessageBoxButtons.OK, MessageBoxIcon.None); }
        public static DialogResult Show(string text, string caption, MessageBoxButtons buttons) { return Show(text, caption, buttons, MessageBoxIcon.None); }
        public static DialogResult Show(string text, string caption, MessageBoxButtons buttons, MessageBoxIcon icon) {
            if (Impl == null) return MessageBox.Show(text, caption, buttons, icon);
            string r = Impl(text, caption, buttons.ToString(), icon.ToString());
            return (DialogResult)System.Enum.Parse(typeof(DialogResult), r);
        }
    }
}
'@
}
function Show-LibDialog([string]$Text, [string]$Caption, [string]$Buttons, [string]$Icon) {
    $warn = $Icon -in @('Warning','Exclamation','Error','Hand','Stop')
    $dlg = New-Object System.Windows.Forms.Form
    $dlg.Text = $(if ($Caption) { $Caption } else { 'Библиотекарь' })
    $dlg.StartPosition = 'CenterParent'; $dlg.FormBorderStyle = 'FixedDialog'; $dlg.MinimizeBox = $false; $dlg.MaximizeBox = $false
    $dlg.BackColor = $script:Paper; $dlg.ShowInTaskbar = $false
    if ($form.Icon) { $dlg.Icon = $form.Icon }
    $font = [System.Drawing.Font]::new('Georgia', 11)
    $textW = 480
    $size = [System.Windows.Forms.TextRenderer]::MeasureText($Text, $font, [System.Drawing.Size]::new($textW, 2000), [System.Windows.Forms.TextFormatFlags]'WordBreak')
    $textH = [math]::Min(560, [math]::Max(120, $size.Height + 10))
    # the Librarian on the left: smiling with a book, or «тсс» when something needs care
    $pic = New-Object System.Windows.Forms.PictureBox; $pic.Location = [System.Drawing.Point]::new(14,16); $pic.Size = [System.Drawing.Size]::new(248,143); $pic.SizeMode = 'Zoom'
    $pic.BackColor = $script:Gold; $pic.Padding = [System.Windows.Forms.Padding]::new(4)
    $picFile = if ($warn) { Join-Path $script:MediaDir 'helper\secret.gif' } else { Join-Path $script:MediaDir 'helper\idle_big.png' }
    if (Test-Path -LiteralPath $picFile) { try { $pic.Image = [System.Drawing.Image]::FromStream((New-Object System.IO.MemoryStream (,[IO.File]::ReadAllBytes($picFile)))) } catch { } }
    $label = New-Object System.Windows.Forms.Label; $label.Font = $font; $label.ForeColor = $script:Ink; $label.Text = $Text
    $label.Location = [System.Drawing.Point]::new(278,18); $label.Size = [System.Drawing.Size]::new($textW, $textH)
    $bottomY = [math]::Max(172, 18 + $textH + 12)
    $result = @{ Value = 'Cancel' }
    $defs = switch ($Buttons) {
        'YesNo' { @(@('Да','Yes'), @('Нет','No')) }
        'YesNoCancel' { @(@('Да','Yes'), @('Нет','No'), @('Отмена','Cancel')) }
        'OKCancel' { @(@('Хорошо','OK'), @('Отмена','Cancel')) }
        'RetryCancel' { @(@('Ещё раз','Retry'), @('Отмена','Cancel')) }
        default { @(,@('Хорошо','OK')) }
    }
    if ($defs.Count -gt 0 -and $defs[0] -is [string]) { $defs = @(,$defs) }   # a single button comes out of «switch» unwrapped
    if ($Buttons -eq 'OK' -or !$Buttons -or $Buttons -eq 'None') { $result.Value = 'OK' }
    $x = 278 + $textW
    $made = @()
    foreach ($d in @($defs)[($defs.Count - 1)..0]) {
        $b = New-Object System.Windows.Forms.Button; $b.Text = $d[0]; $b.Tag = $d[1]
        $b.Size = [System.Drawing.Size]::new(116, 34); $x -= 124; $b.Location = [System.Drawing.Point]::new($x, $bottomY)
        if ($d[1] -in 'Yes','OK','Retry') { Set-WoodButton $b } else { Set-WoodButton $b -Light }
        $b.Add_Click({ $result.Value = [string]$this.Tag; $dlg.Close() })
        $dlg.Controls.Add($b); $made += $b
    }
    $first = $made | Where-Object { $_.Tag -in 'Yes','OK','Retry' } | Select-Object -First 1
    if ($first) { $dlg.AcceptButton = $first }
    $cancelBtn = $made | Where-Object { $_.Tag -in 'Cancel','No' } | Select-Object -First 1
    if ($cancelBtn) { $dlg.CancelButton = $cancelBtn } elseif ($first) { $dlg.CancelButton = $first }
    $dlg.Controls.AddRange(@($pic, $label))
    $dlg.ClientSize = [System.Drawing.Size]::new(278 + $textW + 20, $bottomY + 50)
    $dlg.Add_Paint({ $p = New-Object System.Drawing.Pen $script:Gold, 2; $_.Graphics.DrawRectangle($p, 1, 1, $dlg.ClientSize.Width - 3, $dlg.ClientSize.Height - 3); $p.Dispose() })
    try { [void]$dlg.ShowDialog($form) } finally { if ($pic.Image) { $pic.Image.Dispose() }; $dlg.Dispose() }
    return [string]$result.Value
}
function Enable-SmoothPaint($Control) {
    foreach ($p in 'DoubleBuffered','ResizeRedraw') {
        $prop = $Control.GetType().GetProperty($p, [System.Reflection.BindingFlags]'NonPublic,Instance')
        if ($prop) { $prop.SetValue($Control, $true, $null) }
    }
}
$form.Text = 'Библиотекарь'
$form.BackColor = $script:Parchment
$form.KeyPreview = $true

# --- Header: the castle library from the cartoon, darkened on the left so the words read well ---
$layout.RowStyles[0].Height = 128
foreach ($c in @($rootsLabel, $rootsBox, $addRoot, $removeRoot, $scanButton, $dupeButton, $reportButton, $contentButton, $resetButton, $excelPlanButton)) { $c.Visible = $false }
$bannerPath = Join-Path $PSScriptRoot 'media\banner.jpg'
$script:Banner = if (Test-Path -LiteralPath $bannerPath) { try { [System.Drawing.Image]::FromFile($bannerPath) } catch { $null } } else { $null }
Enable-SmoothPaint $top
$top.BackColor = $script:Wood
$top.Add_Paint({
    $g = $_.Graphics; $w = $top.ClientSize.Width; $h = $top.ClientSize.Height
    if ($w -le 0 -or $h -le 0) { return }
    if ($script:Banner) {
        # «cover»: fill the whole strip, crop what does not fit
        $scale = [math]::Max($w / $script:Banner.Width, $h / $script:Banner.Height)
        $dw = $script:Banner.Width * $scale; $dh = $script:Banner.Height * $scale
        $g.InterpolationMode = 'HighQualityBicubic'
        $g.DrawImage($script:Banner, [single](($w - $dw) / 2), [single](($h - $dh) / 2), [single]$dw, [single]$dh)
    }
    $rect = New-Object System.Drawing.Rectangle 0, 0, $w, $h
    $shade = New-Object System.Drawing.Drawing2D.LinearGradientBrush $rect, ([System.Drawing.Color]::FromArgb(225,38,24,12)), ([System.Drawing.Color]::FromArgb(70,38,24,12)), 0.0
    $g.FillRectangle($shade, $rect); $shade.Dispose()
    $line = New-Object System.Drawing.Pen $script:Gold, 2
    $g.DrawLine($line, 0, $h - 1, $w, $h - 1); $line.Dispose()
})
$titleLabel = New-Object System.Windows.Forms.Label
$titleLabel.Text = 'Библиотекарь'; $titleLabel.Font = $script:TitleFont; $titleLabel.ForeColor = [System.Drawing.Color]::FromArgb(247,214,140)
$titleLabel.BackColor = [System.Drawing.Color]::Transparent; $titleLabel.AutoSize = $true; $titleLabel.Location = [System.Drawing.Point]::new(16,18)
$subtitleLabel = New-Object System.Windows.Forms.Label
$subtitleLabel.Text = 'хранитель ваших файлов'; $subtitleLabel.Font = [System.Drawing.Font]::new('Georgia', 10, [System.Drawing.FontStyle]::Italic)
$subtitleLabel.ForeColor = $script:Cream; $subtitleLabel.BackColor = [System.Drawing.Color]::Transparent; $subtitleLabel.AutoSize = $true; $subtitleLabel.Location = [System.Drawing.Point]::new(20,60)
$keysLabel = New-Object System.Windows.Forms.Label
$keysLabel.Text = 'F1 — подсказка по клавишам'; $keysLabel.Font = [System.Drawing.Font]::new('Georgia', 9, [System.Drawing.FontStyle]::Italic)
$keysLabel.ForeColor = $script:Cream; $keysLabel.BackColor = [System.Drawing.Color]::Transparent; $keysLabel.AutoSize = $true; $keysLabel.Location = [System.Drawing.Point]::new(20,96)
$keysLabel.Cursor = [System.Windows.Forms.Cursors]::Hand
$top.Controls.AddRange(@($titleLabel, $subtitleLabel, $keysLabel))
$searchLabel.BackColor = [System.Drawing.Color]::Transparent; $searchLabel.ForeColor = $script:Cream
$searchLabel.Font = [System.Drawing.Font]::new('Georgia', 10, [System.Drawing.FontStyle]::Italic)
$searchLabel.AutoSize = $false; $searchLabel.AutoEllipsis = $true; $searchLabel.Location = [System.Drawing.Point]::new(300,12)
$searchLabel.Size = [System.Drawing.Size]::new(($form.ClientSize.Width - 300 - 20), 22); $searchLabel.Anchor = 'Top,Left,Right'
$searchBox.Font = [System.Drawing.Font]::new('Segoe UI', 12)
$searchBox.Location = [System.Drawing.Point]::new(300,38); $searchBox.Size = [System.Drawing.Size]::new(($form.ClientSize.Width - 300 - 140), 30); $searchBox.Anchor = 'Top,Left,Right'
$searchBox.BorderStyle = 'FixedSingle'; $searchBox.BackColor = $script:Paper; $searchBox.ForeColor = $script:Ink
$searchButton.Location = [System.Drawing.Point]::new(($form.ClientSize.Width - 128), 37); $searchButton.Size = [System.Drawing.Size]::new(108,31); $searchButton.Anchor = 'Top,Right'
$searchButton.Text = 'Найти'; Set-WoodButton $searchButton
$modeFind.Location = [System.Drawing.Point]::new(300,80); $modeFind.Size = [System.Drawing.Size]::new(92,30); $modeFind.Text = 'Поиск  F3'
$modeTask.Location = [System.Drawing.Point]::new(396,80); $modeTask.Size = [System.Drawing.Size]::new(104,30); $modeTask.Text = 'Задание  F4'
Set-WoodButton $modeFind -Light; Set-WoodButton $modeTask -Light
$micButton.Location = [System.Drawing.Point]::new(512,80); $micButton.Size = [System.Drawing.Size]::new(118,30); $micButton.Text = 'Сказать  F2'
Set-WoodButton $micButton
$voiceCheck.Location = [System.Drawing.Point]::new(644,86); $voiceCheck.BackColor = [System.Drawing.Color]::Transparent; $voiceCheck.ForeColor = $script:Cream

# --- The Librarian's corner inside the picture: the hall, his portrait in a gold frame, the answer on a scroll ---
$hallPath = Join-Path $PSScriptRoot 'media\hall.jpg'
$script:Hall = if (Test-Path -LiteralPath $hallPath) { try { [System.Drawing.Image]::FromFile($hallPath) } catch { $null } } else { $null }
Enable-SmoothPaint $answerPanel
$answerPanel.BackColor = $script:Wood
$answerPanel.Padding = [System.Windows.Forms.Padding]::new(12,10,10,10)
$helperBox.BackColor = [System.Drawing.Color]::FromArgb(60,40,20)
$helperBox.Width = 250
$jobsPanel.BackColor = [System.Drawing.Color]::Transparent
$jobsPanel.Padding = [System.Windows.Forms.Padding]::new(8,2,0,0)
$answerLabel.BackColor = [System.Drawing.Color]::Transparent
$answerLabel.Font = $script:BookFont; $answerLabel.ForeColor = $script:Ink
$answerLabel.Padding = [System.Windows.Forms.Padding]::new(30,16,30,12)
$answerPanel.Add_Paint({
    $g = $_.Graphics; $w = $answerPanel.ClientSize.Width; $h = $answerPanel.ClientSize.Height
    if ($w -le 0 -or $h -le 0) { return }
    $g.SmoothingMode = 'AntiAlias'
    if ($script:Hall) {
        $scale = [math]::Max($w / $script:Hall.Width, $h / $script:Hall.Height)
        $dw = $script:Hall.Width * $scale; $dh = $script:Hall.Height * $scale
        $g.InterpolationMode = 'HighQualityBicubic'
        $g.DrawImage($script:Hall, [single](($w - $dw) / 2), [single](($h - $dh) / 2), [single]$dw, [single]$dh)
        $veil = New-Object System.Drawing.SolidBrush ([System.Drawing.Color]::FromArgb(70,30,18,8)); $g.FillRectangle($veil, 0, 0, $w, $h); $veil.Dispose()
    }
    # gold frame around the portrait
    $hb = $helperBox.Bounds
    $frame = New-Object System.Drawing.Pen $script:Gold, 4
    $g.DrawRectangle($frame, $hb.X - 3, $hb.Y - 3, $hb.Width + 5, $hb.Height + 5); $frame.Dispose()
    # the parchment scroll for the answer
    $lb = $answerLabel.Bounds
    $r = New-Object System.Drawing.RectangleF ($lb.X + 12), ($lb.Y + 2), ($lb.Width - 22), ($lb.Height - 4)
    $path = New-Object System.Drawing.Drawing2D.GraphicsPath; $rad = 18
    $path.AddArc($r.X, $r.Y, $rad, $rad, 180, 90); $path.AddArc($r.Right - $rad, $r.Y, $rad, $rad, 270, 90)
    $path.AddArc($r.Right - $rad, $r.Bottom - $rad, $rad, $rad, 0, 90); $path.AddArc($r.X, $r.Bottom - $rad, $rad, $rad, 90, 90); $path.CloseFigure()
    $fill = New-Object System.Drawing.Drawing2D.LinearGradientBrush $r, ([System.Drawing.Color]::FromArgb(246,250,240,220)), ([System.Drawing.Color]::FromArgb(246,236,218,184)), 90.0
    $g.FillPath($fill, $path); $fill.Dispose()
    $edge = New-Object System.Drawing.Pen $script:Gold, 2; $g.DrawPath($edge, $path); $edge.Dispose(); $path.Dispose()
    # rolled ends of the scroll
    $roll = New-Object System.Drawing.SolidBrush ([System.Drawing.Color]::FromArgb(225,214,190,150))
    $g.FillEllipse($roll, $r.X - 7, $r.Y + 6, 14, $r.Height - 12); $g.FillEllipse($roll, $r.Right - 7, $r.Y + 6, 14, $r.Height - 12); $roll.Dispose()
})
$answerLabel.Add_TextChanged({ $answerPanel.Invalidate() })
$form.WindowState = 'Maximized'
$weeklyButton.Text = 'Еженедельная уборка  F6'; $promptsButton.Text = 'Промпты из канала  F7'; $siteButton.Text = 'Сайт промптов  F8'
foreach ($b in @($weeklyButton, $promptsButton, $siteButton)) { Set-WoodButton $b; $b.Width = 172 }

# --- The file list on parchment with a wooden header ---
$grid.BackColor = $script:Paper; $grid.ForeColor = $script:Ink; $grid.GridLines = $false
$grid.Font = [System.Drawing.Font]::new('Segoe UI', 9.5)
$grid.OwnerDraw = $true
$grid.Add_DrawColumnHeader({
    $e = $_
    $b = New-Object System.Drawing.SolidBrush $script:Wood; $e.Graphics.FillRectangle($b, $e.Bounds); $b.Dispose()
    $p = New-Object System.Drawing.Pen $script:Gold; $e.Graphics.DrawLine($p, $e.Bounds.Right - 1, $e.Bounds.Top + 4, $e.Bounds.Right - 1, $e.Bounds.Bottom - 4); $p.Dispose()
    $r = New-Object System.Drawing.Rectangle ($e.Bounds.X + 6), $e.Bounds.Y, ($e.Bounds.Width - 8), $e.Bounds.Height
    [System.Windows.Forms.TextRenderer]::DrawText($e.Graphics, $e.Header.Text, $grid.Font, $r, $script:Cream, [System.Windows.Forms.TextFormatFlags]'VerticalCenter,Left,EndEllipsis')
})
# The last column stretches to the edge, so the wooden header runs the full width.
$grid.Add_Resize({ if ($grid.Columns.Count) { $grid.Columns[$grid.Columns.Count - 1].Width = -2 } })
$grid.Add_DrawItem({ $_.DrawDefault = $true })
$grid.Add_DrawSubItem({ $_.DrawDefault = $true })
foreach ($c in @($pathPanel, $bottom)) { $c.BackColor = $script:Parchment }
$pathLabel.ForeColor = $script:Ink; $selectedPathBox.BackColor = $script:Paper; $selectedPathBox.ForeColor = $script:Ink; $selectedPathBox.BorderStyle = 'FixedSingle'
Set-WoodButton $copyPathButton -Light
$selectedPathBox.Width = $copyPathButton.Left - $selectedPathBox.Left - 10

# --- Bottom: only «Открыть файл», «Показать в папке» and «Ещё ▾»; the rest lives in the menu ---
foreach ($c in @($planButton, $organizeButton, $undoMoveButton, $cleanupButton)) { $c.Visible = $false }
Set-WoodButton $openButton -Light; Set-WoodButton $folderButton -Light
$moreButton = New-Object System.Windows.Forms.Button
$moreButton.Text = 'Ещё ▾'; $moreButton.Location = [System.Drawing.Point]::new(271,9); $moreButton.Size = [System.Drawing.Size]::new(100,29)
Set-WoodButton $moreButton
$bottom.Controls.Add($moreButton)
$status.Location = [System.Drawing.Point]::new(384,14); $status.Size = [System.Drawing.Size]::new(($form.ClientSize.Width - 384 - 16), 22); $status.ForeColor = $script:Ink
$script:MoreMenu = New-Object System.Windows.Forms.ContextMenuStrip
$script:MoreMenu.BackColor = $script:Paper; $script:MoreMenu.ForeColor = $script:Ink; $script:MoreMenu.Font = [System.Drawing.Font]::new('Segoe UI', 10)
function Invoke-ButtonClick($Button) {
    # A hidden button ignores PerformClick, so its Click handlers are raised directly.
    if (!$Button -or !$Button.Enabled) { return }
    $m = [System.Windows.Forms.Control].GetMethod('OnClick', [System.Reflection.BindingFlags]'NonPublic,Instance')
    $m.Invoke($Button, @([EventArgs]::Empty))
}
function Show-KeysHelp {
    $dlg = New-Object System.Windows.Forms.Form
    $dlg.Text = 'Подсказка'; $dlg.StartPosition = 'CenterParent'; $dlg.Size = [System.Drawing.Size]::new(640,500)
    $dlg.FormBorderStyle = 'FixedDialog'; $dlg.MinimizeBox = $false; $dlg.MaximizeBox = $false; $dlg.BackColor = $script:Paper; $dlg.KeyPreview = $true
    $pic = New-Object System.Windows.Forms.PictureBox; $pic.Location = [System.Drawing.Point]::new(14,60); $pic.Size = [System.Drawing.Size]::new(190,330); $pic.SizeMode = 'Zoom'
    $idle = Join-Path $script:HelperDir 'secret.gif'
    if (Test-Path -LiteralPath $idle) { try { $pic.Image = [System.Drawing.Image]::FromFile($idle) } catch { } }
    $head = New-Object System.Windows.Forms.Label; $head.Text = 'Что я умею по клавишам'; $head.Font = [System.Drawing.Font]::new('Georgia', 16, [System.Drawing.FontStyle]::Bold)
    $head.ForeColor = $script:Wood; $head.AutoSize = $true; $head.Location = [System.Drawing.Point]::new(14,14)
    $body = New-Object System.Windows.Forms.Label; $body.Font = [System.Drawing.Font]::new('Georgia', 11); $body.ForeColor = $script:Ink
    $body.Location = [System.Drawing.Point]::new(216,62); $body.Size = [System.Drawing.Size]::new(400,340)
    $body.Text = "F1`tэта подсказка`r`nF2`tсказать голосом`r`nF3`tрежим «Поиск»`r`nF4`tрежим «Задание»`r`nF5`tобновить каталог`r`nF6`tеженедельная уборка" + $(if ($script:HasPrompts) { "`r`nF7`tпромпты из канала`r`nF8`tсайт промптов" }) + "`r`nF9`tнайти дубликаты`r`nF11`tотменить перенос`r`nF12`tпапки каталога`r`nEsc`tочистить строку`r`nEnter`tнайти / выполнить`r`n`r`nВсё остальное — в кнопке «Ещё ▾» внизу."
    $ok = New-Object System.Windows.Forms.Button; $ok.Text = 'Понятно'; $ok.Size = [System.Drawing.Size]::new(120,32); $ok.Location = [System.Drawing.Point]::new(496,412); $ok.DialogResult = 'OK'
    Set-WoodButton $ok
    $dlg.Controls.AddRange(@($pic,$head,$body,$ok)); $dlg.AcceptButton = $ok; $dlg.CancelButton = $ok
    $dlg.Add_KeyDown({ if ($_.KeyCode -eq 'F1') { $dlg.Close() } })
    try { [void]$dlg.ShowDialog($form) } finally { if ($pic.Image) { $pic.Image.Dispose() }; $dlg.Dispose() }
}
function Show-RootsDialog {
    # The folder list moves into this window for a moment and goes back afterwards.
    $dlg = New-Object System.Windows.Forms.Form
    $dlg.Text = 'Папки каталога'; $dlg.StartPosition = 'CenterParent'; $dlg.Size = [System.Drawing.Size]::new(720,330)
    $dlg.FormBorderStyle = 'FixedDialog'; $dlg.MinimizeBox = $false; $dlg.MaximizeBox = $false; $dlg.BackColor = $script:Paper
    $l = New-Object System.Windows.Forms.Label; $l.Text = 'Где Библиотекарь ищет файлы (после изменений нажмите F5 — «Обновить каталог»):'; $l.AutoSize = $true; $l.Location = [System.Drawing.Point]::new(12,12); $l.ForeColor = $script:Ink
    $close = New-Object System.Windows.Forms.Button; $close.Text = 'Готово'; $close.Size = [System.Drawing.Size]::new(110,32); $close.Location = [System.Drawing.Point]::new(582,246); $close.DialogResult = 'OK'
    Set-WoodButton $close
    $oldParent = $rootsBox.Parent
    $saved = @{ BoxLoc = $rootsBox.Location; BoxSize = $rootsBox.Size; AddLoc = $addRoot.Location; RemLoc = $removeRoot.Location }
    $rootsBox.Location = [System.Drawing.Point]::new(12,38); $rootsBox.Size = [System.Drawing.Size]::new(540,200)
    $addRoot.Location = [System.Drawing.Point]::new(562,38); $removeRoot.Location = [System.Drawing.Point]::new(562,72)
    foreach ($c in @($rootsBox, $addRoot, $removeRoot)) { $dlg.Controls.Add($c); $c.Visible = $true }
    Set-WoodButton $addRoot -Light; Set-WoodButton $removeRoot -Light
    $dlg.Controls.AddRange(@($l, $close)); $dlg.AcceptButton = $close
    try { [void]$dlg.ShowDialog($form) } finally {
        foreach ($c in @($rootsBox, $addRoot, $removeRoot)) { $oldParent.Controls.Add($c); $c.Visible = $false }
        $rootsBox.Location = $saved.BoxLoc; $rootsBox.Size = $saved.BoxSize; $addRoot.Location = $saved.AddLoc; $removeRoot.Location = $saved.RemLoc
        $dlg.Dispose()
    }
}
function Add-MoreItem([string]$Text, [string]$Keys, $Button) {
    $it = New-Object System.Windows.Forms.ToolStripMenuItem $Text
    $it.ShortcutKeyDisplayString = $Keys; $it.Tag = $Button
    $it.Add_Click({
        $tag = $this.Tag
        if ($tag -eq 'help') { Show-KeysHelp } elseif ($tag -eq 'roots') { Show-RootsDialog } elseif ($tag -eq 'demo') { Start-SceneDemo } elseif ($tag -eq 'schedule') { Set-WeeklySchedule } else { Invoke-ButtonClick $tag }
    })
    [void]$script:MoreMenu.Items.Add($it)
}
Add-MoreItem 'Подсказка по клавишам' 'F1' $null
Add-MoreItem 'Показать все сцены' '' $null
Add-MoreItem 'Обновить каталог' 'F5' $scanButton
Add-MoreItem 'Найти дубликаты' 'F9' $dupeButton
Add-MoreItem 'Убрать точные дубли' '' $cleanupButton
Add-MoreItem 'Отчёт о дубликатах' '' $reportButton
Add-MoreItem 'Читать содержание документов' '' $contentButton
[void]$script:MoreMenu.Items.Add((New-Object System.Windows.Forms.ToolStripSeparator))
Add-MoreItem 'План раскладки Рабочего стола' '' $planButton
Add-MoreItem 'Открыть план в Excel' '' $excelPlanButton
Add-MoreItem 'Выполнить план…' '' $organizeButton
Add-MoreItem 'Отменить перенос' 'F11' $undoMoveButton
[void]$script:MoreMenu.Items.Add((New-Object System.Windows.Forms.ToolStripSeparator))
Add-MoreItem 'Папки каталога…' 'F12' $null
$script:MoreMenu.Items[0].Tag = 'help'; $script:MoreMenu.Items[1].Tag = 'demo'; $script:MoreMenu.Items[$script:MoreMenu.Items.Count - 1].Tag = 'roots'
Add-MoreItem 'Уборка по воскресеньям…' '' $null
$script:MoreMenu.Items[$script:MoreMenu.Items.Count - 1].Tag = 'schedule'
function Set-WeeklySchedule {
    # Windows Task Scheduler opens the Librarian straight into the weekly cleanup on Sundays at 11:00
    # (or as soon as the computer is switched on, if it was off at that time).
    $name = 'Библиотекарь — еженедельная уборка'
    $task = Get-ScheduledTask -TaskName $name -ErrorAction SilentlyContinue
    if ($task) {
        if ([LibrarianApp.LibBox]::Show("Уборка по воскресеньям уже включена: в 11:00 я сам прихожу с метлой.`r`n`r`nВыключить?", 'Уборка по воскресеньям', 'YesNo', 'Question') -eq 'Yes') {
            Unregister-ScheduledTask -TaskName $name -Confirm:$false
            $answerLabel.Text = 'Хорошо, по воскресеньям больше сам не прихожу. Уборка — табличкой «Еженедельная уборка» или F6.'
        }
        return
    }
    if ([LibrarianApp.LibBox]::Show("Приходить с уборкой каждое воскресенье в 11:00?`r`n`r`nЯ открою окно и покажу, что нашёл в «Загрузках» и на Рабочем столе. Без вашего «да» ничего не трогаю, кроме старых установщиков и точных копий — они уходят в Корзину.", 'Уборка по воскресеньям', 'YesNo', 'Question') -ne 'Yes') { return }
    try {
        $self = $PSCommandPath; if (!$self) { $self = $MyInvocation.ScriptName }
        $action = New-ScheduledTaskAction -Execute (Join-Path $PSHOME 'powershell.exe') -Argument ('-NoProfile -STA -ExecutionPolicy Bypass -WindowStyle Hidden -File "' + $self + '" -WeeklyCleanup')
        $trigger = New-ScheduledTaskTrigger -Weekly -DaysOfWeek Sunday -At '11:00'
        $taskSettings = New-ScheduledTaskSettingsSet -StartWhenAvailable -AllowStartIfOnBatteries -DontStopIfGoingOnBatteries
        Register-ScheduledTask -TaskName $name -Action $action -Trigger $trigger -Settings $taskSettings -Description 'Еженедельная уборка «Загрузок» и Рабочего стола' -ErrorAction Stop | Out-Null
        $answerLabel.Text = 'Готово: каждое воскресенье в 11:00 прихожу с уборкой. Выключить — «Ещё ▾» → «Уборка по воскресеньям».'
    } catch { $answerLabel.Text = 'Не получилось включить расписание: ' + $_.Exception.Message }
}
$script:MoreMenu.Add_Opening({ foreach ($it in $script:MoreMenu.Items) { if ($it.Tag -is [System.Windows.Forms.Button]) { $it.Enabled = $it.Tag.Enabled } } })
$moreButton.Add_Click({ $script:MoreMenu.Show($moreButton, 0, -($script:MoreMenu.PreferredSize.Height)) })
$keysLabel.Add_Click({ Show-KeysHelp })
$form.Add_KeyDown({
    $k = $_.KeyCode; $done = $true
    switch ($k) {
        'F1' { Show-KeysHelp }
        'F2' { $micButton.PerformClick() }
        'F3' { $modeFind.Checked = $true; [void]$searchBox.Focus() }
        'F4' { $modeTask.Checked = $true }
        'F5' { Invoke-ButtonClick $scanButton }
        'F6' { $weeklyButton.PerformClick() }
        'F7' { if ($script:HasPrompts) { $promptsButton.PerformClick() } }
        'F8' { if ($script:HasPrompts) { $siteButton.PerformClick() } }
        'F9' { Invoke-ButtonClick $dupeButton }
        'F11' { Invoke-ButtonClick $undoMoveButton }
        'F12' { Show-RootsDialog }
        'L' { if ($_.Control) { Switch-FileList } else { $done = $false } }
        'Escape' { if ($searchBox.Text) { $searchBox.Clear(); [void]$searchBox.Focus() } else { $done = $false } }
        default { $done = $false }
    }
    if ($done) { $_.Handled = $true; $_.SuppressKeyPress = $true }
})

# ================= Одна большая картина: всё управление прямо на зале библиотеки =================
# The window is the hall; the file table slides out from below only when asked (button or Ctrl+L).
$stagePath = Join-Path $PSScriptRoot 'media\stage.png'
$script:StageImage = if (Test-Path -LiteralPath $stagePath) { try { [System.Drawing.Image]::FromStream((New-Object System.IO.MemoryStream (,[IO.File]::ReadAllBytes($stagePath)))) } catch { $null } } else { $null }
$stage = New-Object System.Windows.Forms.Panel
$stage.Dock = 'Fill'; $stage.BackColor = $script:Wood
Enable-SmoothPaint $stage
$listPanel = New-Object System.Windows.Forms.Panel
$listPanel.Dock = 'Bottom'; $listPanel.Height = 380; $listPanel.Visible = $false; $listPanel.BackColor = $script:Parchment
$listSplit = New-Object System.Windows.Forms.Splitter
$listSplit.Dock = 'Bottom'; $listSplit.Height = 5; $listSplit.BackColor = $script:Gold; $listSplit.Visible = $false
# The old table layout keeps only the list, the path and the bottom buttons.
$form.Controls.Remove($layout)
$layout.RowStyles[0].Height = 0; $layout.RowStyles[1].Height = 0
$top.Visible = $false; $answerPanel.Visible = $false
$listPanel.Controls.Add($layout)
foreach ($c in @($titleLabel, $subtitleLabel, $keysLabel, $searchLabel, $searchBox, $searchButton, $modeFind, $modeTask, $micButton, $voiceCheck,
                  $helperBox, $answerLabel, $weeklyButton, $promptsButton, $siteButton)) { $stage.Controls.Add($c) }
$helperBox.Dock = 'None'; $answerLabel.Dock = 'None'; $helperBox.Anchor = 'Top,Left'; $answerLabel.Anchor = 'Top,Left'
foreach ($c in @($searchLabel, $searchBox, $searchButton)) { $c.Anchor = 'Top,Left' }
$answerLabel.Font = [System.Drawing.Font]::new('Georgia', 13); $answerLabel.Padding = [System.Windows.Forms.Padding]::new(34,22,34,16)
$helperBox.BackColor = [System.Drawing.Color]::FromArgb(40,26,12)
$openFound = New-Object System.Windows.Forms.Button; $openFound.Text = 'Открыть найденное'; Set-WoodButton $openFound
$showFound = New-Object System.Windows.Forms.Button; $showFound.Text = 'Показать в папке'; Set-WoodButton $showFound
$listToggle = New-Object System.Windows.Forms.Button; $listToggle.Text = 'Список файлов  Ctrl+L'; Set-WoodButton $listToggle -Light
$stageMore = New-Object System.Windows.Forms.Button; $stageMore.Text = 'Ещё ▾'; Set-WoodButton $stageMore
$stage.Controls.AddRange(@($openFound, $showFound, $listToggle, $stageMore))
$openFound.Add_Click({ Open-SelectedFile })
$showFound.Add_Click({ Show-SelectedFileInFolder })
$stageMore.Add_Click({ $script:MoreMenu.Show($stageMore, 0, $stageMore.Height) })
function Switch-FileList {
    $show = !$listPanel.Visible
    $listPanel.Visible = $show; $listSplit.Visible = $show
    $listToggle.Text = if ($show) { 'Скрыть список  Ctrl+L' } else { 'Список файлов  Ctrl+L' }
    if ($show) { $listPanel.Height = [int]($form.ClientSize.Height * 0.45) }
}
$listToggle.Add_Click({ Switch-FileList })
$form.Controls.Add($stage); $form.Controls.Add($listSplit); $form.Controls.Add($listPanel)
$stage.BringToFront()
$script:StageBox = @{ Scroll = [System.Drawing.Rectangle]::Empty }
function Set-StageLayout {
    # Everything is placed by the size of the picture, so it fits any screen.
    $w = $stage.ClientSize.Width; $h = $stage.ClientSize.Height
    if ($w -lt 300 -or $h -lt 200) { return }
    $m = 24
    $titleLabel.Location = [System.Drawing.Point]::new($m, 18); $subtitleLabel.Location = [System.Drawing.Point]::new($m + 4, 62); $keysLabel.Location = [System.Drawing.Point]::new($m + 4, 92)
    $x0 = 320
    $searchLabel.Location = [System.Drawing.Point]::new($x0, 14); $searchLabel.Size = [System.Drawing.Size]::new([math]::Max(200, $w - $x0 - $m), 22)
    $searchButton.Size = [System.Drawing.Size]::new(118, 32); $searchButton.Location = [System.Drawing.Point]::new($w - $m - 118, 40)
    $searchBox.Location = [System.Drawing.Point]::new($x0, 41); $searchBox.Width = [math]::Max(200, $searchButton.Left - $x0 - 10)
    $modeFind.Location = [System.Drawing.Point]::new($x0, 84); $modeTask.Location = [System.Drawing.Point]::new($x0 + 98, 84)
    $micButton.Location = [System.Drawing.Point]::new($x0 + 214, 84); $voiceCheck.Location = [System.Drawing.Point]::new($x0 + 348, 90)
    # body: portrait — scroll — wooden plaques
    $bodyTop = 140; $btnH = 40; $bottomBtns = 56
    $ph = [math]::Max(170, [math]::Min(460, $h - $bodyTop - $bottomBtns - 20)); $pw = [int]($ph * 16 / 9)
    $plaqueW = 220; $px = $w - $m - $plaqueW
    if ($pw -gt ($w * 0.42)) { $pw = [int]($w * 0.42); $ph = [int]($pw * 9 / 16) }
    $helperBox.Location = [System.Drawing.Point]::new($m + 4, $bodyTop + 4); $helperBox.Size = [System.Drawing.Size]::new($pw, $ph)
    $sx = $m + $pw + 30; $sw = [math]::Max(220, $px - 24 - $sx)
    $answerLabel.Location = [System.Drawing.Point]::new($sx, $bodyTop); $answerLabel.Size = [System.Drawing.Size]::new($sw, $ph + 8)
    $script:StageBox.Scroll = New-Object System.Drawing.Rectangle $sx, $bodyTop, $sw, ($ph + 8)
    $i = 0
    foreach ($b in (Get-Plaques)) { $b.Size = [System.Drawing.Size]::new($plaqueW, $btnH + 4); $b.Location = [System.Drawing.Point]::new($px, $bodyTop + $i * ($btnH + 14)); $i++ }
    $by = $bodyTop + $ph + 22
    $openFound.Size = [System.Drawing.Size]::new(180, $btnH); $openFound.Location = [System.Drawing.Point]::new($sx + 14, $by)
    $showFound.Size = [System.Drawing.Size]::new(170, $btnH); $showFound.Location = [System.Drawing.Point]::new($sx + 204, $by)
    $listToggle.Size = [System.Drawing.Size]::new(200, $btnH); $listToggle.Location = [System.Drawing.Point]::new($sx + 384, $by)
    $stage.Invalidate()
}
$stage.Add_Resize({ Set-StageLayout })
$stage.Add_Paint({
    $g = $_.Graphics; $w = $stage.ClientSize.Width; $h = $stage.ClientSize.Height
    if ($w -le 0 -or $h -le 0) { return }
    if ($script:FullStage -and $script:Scene) { return }   # the live scene paints the whole window itself
    $g.SmoothingMode = 'AntiAlias'; $g.InterpolationMode = 'HighQualityBicubic'
    if ($script:StageImage) {
        $scale = [math]::Max($w / $script:StageImage.Width, $h / $script:StageImage.Height)
        $dw = $script:StageImage.Width * $scale; $dh = $script:StageImage.Height * $scale
        $g.DrawImage($script:StageImage, [single](($w - $dw) / 2), [single](($h - $dh) / 2), [single]$dw, [single]$dh)
    }
    # soft darkening: strong under the title and search, light over the hall
    $all = New-Object System.Drawing.SolidBrush ([System.Drawing.Color]::FromArgb(70,25,15,6)); $g.FillRectangle($all, 0, 0, $w, $h); $all.Dispose()
    $topRect = New-Object System.Drawing.Rectangle 0, 0, $w, 132
    $band = New-Object System.Drawing.Drawing2D.LinearGradientBrush $topRect, ([System.Drawing.Color]::FromArgb(215,30,18,8)), ([System.Drawing.Color]::FromArgb(150,30,18,8)), 0.0
    $g.FillRectangle($band, $topRect); $band.Dispose()
    $goldPen = New-Object System.Drawing.Pen $script:Gold, 2; $g.DrawLine($goldPen, 0, 131, $w, 131); $goldPen.Dispose()
    # gold frame of the portrait
    $hb = $helperBox.Bounds
    $frame = New-Object System.Drawing.Pen $script:Gold, 5
    $g.DrawRectangle($frame, $hb.X - 4, $hb.Y - 4, $hb.Width + 7, $hb.Height + 7); $frame.Dispose()
    $inner = New-Object System.Drawing.Pen ([System.Drawing.Color]::FromArgb(120,70,30)), 1
    $g.DrawRectangle($inner, $hb.X - 8, $hb.Y - 8, $hb.Width + 15, $hb.Height + 15); $inner.Dispose()
    # the parchment scroll
    $sb = $script:StageBox.Scroll
    if ($sb.Width -gt 0) {
        $r = New-Object System.Drawing.RectangleF ($sb.X + 10), ($sb.Y + 2), ($sb.Width - 20), ($sb.Height - 4)
        $path = New-Object System.Drawing.Drawing2D.GraphicsPath; $rad = 22
        $path.AddArc($r.X, $r.Y, $rad, $rad, 180, 90); $path.AddArc($r.Right - $rad, $r.Y, $rad, $rad, 270, 90)
        $path.AddArc($r.Right - $rad, $r.Bottom - $rad, $rad, $rad, 0, 90); $path.AddArc($r.X, $r.Bottom - $rad, $rad, $rad, 90, 90); $path.CloseFigure()
        $fill = New-Object System.Drawing.Drawing2D.LinearGradientBrush $r, ([System.Drawing.Color]::FromArgb(242,251,242,222)), ([System.Drawing.Color]::FromArgb(242,236,218,182)), 90.0
        $g.FillPath($fill, $path); $fill.Dispose()
        $edge = New-Object System.Drawing.Pen $script:Gold, 2; $g.DrawPath($edge, $path); $edge.Dispose(); $path.Dispose()
        $roll = New-Object System.Drawing.SolidBrush ([System.Drawing.Color]::FromArgb(230,206,180,136))
        $g.FillEllipse($roll, $r.X - 8, $r.Y + 6, 16, $r.Height - 12); $g.FillEllipse($roll, $r.Right - 8, $r.Y + 6, 16, $r.Height - 12); $roll.Dispose()
    }
})
$answerLabel.Add_TextChanged({ $stage.Invalidate($script:StageBox.Scroll) })
$form.Add_Shown({ Set-StageLayout })

# ================= Живая сцена: мультик идёт во всё окно, без рамки =================
# media\helper_full\*.gif (960x540) play behind everything; between scenes — the still idle.jpg.
# Remove the folder helper_full (or restore the backup «картина в раме») to go back to the framed portrait.
$script:FullDir = Join-Path $PSScriptRoot 'media\helper_full'
$script:FullStage = Test-Path -LiteralPath (Join-Path $script:FullDir 'idle.jpg')
$script:Scene = $null; $script:SceneFrames = 0; $script:SceneIndex = 0
$script:SceneTimer = New-Object System.Windows.Forms.Timer
$script:SceneTimer.Interval = 25   # checked often; a new frame is drawn only when it is due
$script:SceneFrameMs = 1000.0 / 12  # the scenes were made at 12 frames per second
$script:SceneClock = New-Object System.Diagnostics.Stopwatch
$script:SceneBaseMs = 0.0; $script:SceneUsesVoice = $false; $script:LastVoiceMs = 0.0
$script:SceneTimer.Add_Tick({
    if (!$script:Scene -or $script:SceneFrames -le 1) { $script:SceneTimer.Stop(); return }
    # Lips follow the voice: while Arthur speaks, the frame is taken from the voice position itself,
    # so a slow timer or a late start of the sound can never pull them apart.
    $ms = 0.0
    if ($script:SceneUsesVoice -and $script:VoicePlayer) {
        $pos = $script:VoicePlayer.Position.TotalMilliseconds
        $dur = if ($script:VoicePlayer.NaturalDuration.HasTimeSpan) { $script:VoicePlayer.NaturalDuration.TimeSpan.TotalMilliseconds } else { 0 }
        if ($pos -gt 0 -and ($dur -eq 0 -or $pos -lt $dur - 40)) { $ms = $pos; $script:LastVoiceMs = $pos; $script:SceneClock.Restart() }
        elseif ($pos -le 0 -and $script:SceneClock.ElapsedMilliseconds -lt 900) { $ms = 0 }   # the voice is still opening: wait on the first frame
        else { $script:SceneUsesVoice = $false; $script:SceneBaseMs = $script:LastVoiceMs; $script:SceneClock.Restart(); $ms = $script:SceneBaseMs }
    } else { $ms = $script:SceneBaseMs + $script:SceneClock.ElapsedMilliseconds }
    $idx = [int][math]::Floor($ms / $script:SceneFrameMs) % $script:SceneFrames
    if ($idx -ne $script:SceneIndex) {
        $script:SceneIndex = $idx
        try { [void]$script:Scene.SelectActiveFrame([System.Drawing.Imaging.FrameDimension]::Time, $idx) } catch { }
        $stage.Invalidate()
    }
})
function Set-SceneImage([string]$Path) {
    $script:SceneTimer.Stop()
    try {
        $img = [System.Drawing.Image]::FromStream((New-Object System.IO.MemoryStream (,[IO.File]::ReadAllBytes($Path))))
        $old = $script:Scene; $script:Scene = $img; if ($old) { $old.Dispose() }
        $script:SceneFrames = if ($Path -like '*.gif') { try { $img.GetFrameCount([System.Drawing.Imaging.FrameDimension]::Time) } catch { 1 } } else { 1 }
        $script:SceneIndex = 0; $script:SceneBaseMs = 0.0; $script:SceneUsesVoice = $false; $script:LastVoiceMs = 0.0
        $script:SceneClock.Restart()
        if ($script:SceneFrames -gt 1) { $script:SceneTimer.Start() }
    } catch { }
    $stage.Invalidate()
}
if ($script:FullStage) {
    $helperBox.Visible = $false
    # Same names as the framed version: Show-Helper / Show-HelperIdle now drive the whole window.
    # Defined at the very end of the script (see «FullOverrides»), so they replace the framed versions.
    $script:FullOverrides = {
    function Set-HelperImage([string]$Path) {
        $name = [IO.Path]::GetFileName($Path)
        $full = Join-Path $script:FullDir $name
        if ($name -like 'idle*') { $full = Join-Path $script:FullDir 'idle.jpg' }
        if (Test-Path -LiteralPath $full) { Set-SceneImage $full }
    }
    function Set-StageLayout {
        # No top band: the whole conversation lives in one big scroll on the left — title, the Librarian's
        # answers, and at the bottom the line where you write or speak. He lives in the middle of the picture.
        $w = $stage.ClientSize.Width; $h = $stage.ClientSize.Height
        if ($w -lt 300 -or $h -lt 200) { return }
        $m = 20; $btnH = 40
        $sw = [math]::Max(400, [math]::Min(620, [int]($w * 0.31)))
        $sx = $m; $sy = $m; $sh = $h - 2 * $m - (2 * $btnH + 24)
        $script:StageBox.Scroll = New-Object System.Drawing.Rectangle $sx, $sy, $sw, $sh
        $ix = $sx + 36; $iw = $sw - 72; $bottom = $sy + $sh
        $titleLabel.ForeColor = $script:Wood; $subtitleLabel.ForeColor = [System.Drawing.Color]::FromArgb(130,95,55); $keysLabel.ForeColor = [System.Drawing.Color]::FromArgb(130,95,55)
        $titleLabel.Location = [System.Drawing.Point]::new($ix - 4, $sy + 18)
        $subtitleLabel.Location = [System.Drawing.Point]::new($ix, $sy + 58)
        $keysLabel.Location = [System.Drawing.Point]::new($ix + $iw - $keysLabel.Width, $sy + 64)
        # the input area at the bottom of the scroll
        $rowY = $bottom - 58; $inY = $rowY - 44; $hintY = $inY - 26
        $script:StageBox.InputTop = $hintY - 10
        $answerLabel.Location = [System.Drawing.Point]::new($sx + 18, $sy + 88); $answerLabel.Size = [System.Drawing.Size]::new($sw - 36, [math]::Max(60, $hintY - 14 - ($sy + 88)))
        $answerLabel.Padding = [System.Windows.Forms.Padding]::new(18,8,18,8)
        $searchLabel.ForeColor = [System.Drawing.Color]::FromArgb(120,88,52); $searchLabel.Font = [System.Drawing.Font]::new('Georgia', 9.5, [System.Drawing.FontStyle]::Italic)
        $searchLabel.Location = [System.Drawing.Point]::new($ix, $hintY); $searchLabel.Size = [System.Drawing.Size]::new($iw, 22)
        $searchBox.Location = [System.Drawing.Point]::new($ix, $inY); $searchBox.Width = $iw
        $cw = [int](($iw - 18) / 4)
        $modeFind.Location = [System.Drawing.Point]::new($ix, $rowY); $modeFind.Size = [System.Drawing.Size]::new($cw, 34)
        $modeTask.Location = [System.Drawing.Point]::new($ix + ($cw + 6), $rowY); $modeTask.Size = [System.Drawing.Size]::new($cw, 34)
        $micButton.Location = [System.Drawing.Point]::new($ix + 2 * ($cw + 6), $rowY); $micButton.Size = [System.Drawing.Size]::new($cw, 34)
        $searchButton.Location = [System.Drawing.Point]::new($ix + 3 * ($cw + 6), $rowY); $searchButton.Size = [System.Drawing.Size]::new($iw - 3 * ($cw + 6), 34)
        $voiceCheck.ForeColor = $script:Cream; $voiceCheck.BackColor = [System.Drawing.Color]::Transparent
        # under the scroll: what to do with the found file
        $by = $bottom + 12; $bw = [int](($sw - 20) / 2)
        $openFound.Size = [System.Drawing.Size]::new($bw, $btnH); $openFound.Location = [System.Drawing.Point]::new($sx + 10, $by)
        $showFound.Size = [System.Drawing.Size]::new($bw, $btnH); $showFound.Location = [System.Drawing.Point]::new($sx + 10 + $bw, $by)
        $listToggle.Size = [System.Drawing.Size]::new($sw - 20 - 150, $btnH); $listToggle.Location = [System.Drawing.Point]::new($sx + 10, $by + $btnH + 8)
        $voiceCheck.Location = [System.Drawing.Point]::new($listToggle.Right + 14, $by + $btnH + 18)
        # wooden plaques on the right
        $plaqueW = 230; $px = $w - $m - $plaqueW; $i = 0
        foreach ($b in (Get-Plaques)) { $b.Size = [System.Drawing.Size]::new($plaqueW, $btnH + 4); $b.Location = [System.Drawing.Point]::new($px, $m + $i * ($btnH + 14)); $i++ }
        # the see-through file list: lower right part of the picture
        $ox = $sx + $sw + 24; $oy = [int]($h * 0.42)
        $script:OverlayRect = New-Object System.Drawing.Rectangle $ox, $oy, ([math]::Max(300, $w - $ox - $m)), ([math]::Max(140, $h - $oy - $m))
        $stage.Invalidate()
    }    function Switch-FileList {
        # In the live scene the list is written on the picture — the picture itself never moves.
        $script:ListOverlay = !$script:ListOverlay
        $listToggle.Text = if ($script:ListOverlay) { 'Скрыть список  Ctrl+L' } else { 'Список файлов  Ctrl+L' }
        $sel = if ($grid.SelectedIndices.Count) { $grid.SelectedIndices[0] } else { 0 }
        $script:ListTopRow = [math]::Max(0, $sel - 2)
        $stage.Invalidate()
    }
    }
    # ---- The file list lies on the picture itself: words on a see-through dark band, the scene keeps playing ----
    $script:ListOverlay = $false; $script:ListTopRow = 0; $script:OverlayRect = [System.Drawing.Rectangle]::Empty
    $script:OverlayRowH = 26; $script:PlaceCache = @{}
    $script:ListFont = [System.Drawing.Font]::new('Segoe UI', 10.5); $script:ListHeadFont = [System.Drawing.Font]::new('Georgia', 10.5, [System.Drawing.FontStyle]::Bold)
    if (-not ('LibrarianApp.WheelFilter' -as [type])) {
        Add-Type -ReferencedAssemblies System.Windows.Forms, System.Drawing -TypeDefinition @'
namespace LibrarianApp {
    public class WheelFilter : System.Windows.Forms.IMessageFilter {
        public static int Delta; public static System.Drawing.Point At;
        public event System.EventHandler Wheel;
        public bool PreFilterMessage(ref System.Windows.Forms.Message m) {
            if (m.Msg != 0x020A) return false;
            long w = m.WParam.ToInt64(); Delta = (short)((w >> 16) & 0xFFFF);
            long l = m.LParam.ToInt64(); At = new System.Drawing.Point((short)(l & 0xFFFF), (short)((l >> 16) & 0xFFFF));
            if (Wheel != null) Wheel(this, System.EventArgs.Empty);
            return false;
        }
    }
}
'@
    }
    $script:Wheel = New-Object LibrarianApp.WheelFilter
    [System.Windows.Forms.Application]::AddMessageFilter($script:Wheel)
    function Get-OverlayVisibleRows { if ($script:OverlayRect.Height -le 0) { return 0 }; return [math]::Max(1, [int][math]::Floor(($script:OverlayRect.Height - 34 - 26) / $script:OverlayRowH)) }
    function Move-OverlaySelection([int]$Delta) {
        $n = $grid.Items.Count; if ($n -eq 0) { return }
        $cur = if ($grid.SelectedIndices.Count) { $grid.SelectedIndices[0] } else { -1 }
        $i = [math]::Max(0, [math]::Min($n - 1, $cur + $Delta))
        foreach ($si in @($grid.SelectedItems)) { $si.Selected = $false }; $grid.Items[$i].Selected = $true; $grid.Items[$i].Focused = $true; $script:UserPick = [string]$grid.Items[$i].Tag
        $vis = Get-OverlayVisibleRows
        if ($i -lt $script:ListTopRow) { $script:ListTopRow = $i } elseif ($i -ge $script:ListTopRow + $vis) { $script:ListTopRow = $i - $vis + 1 }
        $stage.Invalidate($script:OverlayRect)
    }
    $script:Wheel.Add_Wheel({
        if (!$script:ListOverlay) { return }
        $p = $stage.PointToClient([System.Windows.Forms.Cursor]::Position)
        if (!$script:OverlayRect.Contains($p)) { return }
        $rows = 3 * [math]::Sign(-[LibrarianApp.WheelFilter]::Delta)
        $script:ListTopRow = [math]::Max(0, [math]::Min([math]::Max(0, $grid.Items.Count - (Get-OverlayVisibleRows)), $script:ListTopRow + $rows))
        $stage.Invalidate($script:OverlayRect)
    })
    $stage.Add_MouseDown({
        if (!$script:ListOverlay -or !$script:OverlayRect.Contains($_.Location)) { return }
        $row = [int][math]::Floor(($_.Y - $script:OverlayRect.Y - 32) / $script:OverlayRowH)
        if ($row -lt 0) { return }
        $i = $script:ListTopRow + $row
        if ($i -ge $grid.Items.Count) { return }
        foreach ($si in @($grid.SelectedItems)) { $si.Selected = $false }; $grid.Items[$i].Selected = $true; $grid.Items[$i].Focused = $true; $script:UserPick = [string]$grid.Items[$i].Tag
        $stage.Invalidate($script:OverlayRect)
    })
    $stage.Add_DoubleClick({ if ($script:ListOverlay -and $grid.SelectedItems.Count) { Open-SelectedFile } })
    $form.Add_KeyDown({
        if (!$script:ListOverlay) { return }
        switch ($_.KeyCode) {
            'Down' { Move-OverlaySelection 1; $_.Handled = $true }
            'Up' { Move-OverlaySelection -1; $_.Handled = $true }
            'PageDown' { Move-OverlaySelection (Get-OverlayVisibleRows); $_.Handled = $true }
            'PageUp' { Move-OverlaySelection (-(Get-OverlayVisibleRows)); $_.Handled = $true }
        }
    })
    $grid.Add_SelectedIndexChanged({ if ($script:ListOverlay) { $stage.Invalidate($script:OverlayRect) } })
    $stage.Add_Paint({
        # drawn after the static hall: the live scene covers it completely
        if (!$script:Scene) { return }
        $g = $_.Graphics; $w = $stage.ClientSize.Width; $h = $stage.ClientSize.Height
        $g.InterpolationMode = 'Bilinear'
        $scale = [math]::Max($w / $script:Scene.Width, $h / $script:Scene.Height)
        $dw = $script:Scene.Width * $scale; $dh = $script:Scene.Height * $scale
        $g.DrawImage($script:Scene, [single](($w - $dw) / 2), [single](($h - $dh) / 2), [single]$dw, [single]$dh)
        $sb = $script:StageBox.Scroll
        if ($sb.Width -gt 0) {
            $g.SmoothingMode = 'AntiAlias'
            $r = New-Object System.Drawing.RectangleF ($sb.X + 10), ($sb.Y + 2), ($sb.Width - 20), ($sb.Height - 4)
            $path = New-Object System.Drawing.Drawing2D.GraphicsPath; $rad = 22
            $path.AddArc($r.X, $r.Y, $rad, $rad, 180, 90); $path.AddArc($r.Right - $rad, $r.Y, $rad, $rad, 270, 90)
            $path.AddArc($r.Right - $rad, $r.Bottom - $rad, $rad, $rad, 0, 90); $path.AddArc($r.X, $r.Bottom - $rad, $rad, $rad, 90, 90); $path.CloseFigure()
            $fill = New-Object System.Drawing.Drawing2D.LinearGradientBrush $r, ([System.Drawing.Color]::FromArgb(236,251,242,222)), ([System.Drawing.Color]::FromArgb(236,236,218,182)), 90.0
            $g.FillPath($fill, $path); $fill.Dispose()
            $edge = New-Object System.Drawing.Pen $script:Gold, 2; $g.DrawPath($edge, $path); $edge.Dispose(); $path.Dispose()
            $roll = New-Object System.Drawing.SolidBrush ([System.Drawing.Color]::FromArgb(230,206,180,136))
            $g.FillEllipse($roll, $r.X - 8, $r.Y + 6, 16, $r.Height - 12); $g.FillEllipse($roll, $r.Right - 8, $r.Y + 6, 16, $r.Height - 12); $roll.Dispose()
            # gold lines: under the title and above the place where you write
            $sep = New-Object System.Drawing.Pen ([System.Drawing.Color]::FromArgb(170,196,150,72)), 1
            $g.DrawLine($sep, $r.X + 26, $r.Y + 82, $r.Right - 26, $r.Y + 82)
            if ($script:StageBox.InputTop) { $g.DrawLine($sep, $r.X + 26, $script:StageBox.InputTop, $r.Right - 26, $script:StageBox.InputTop) }
            $sep.Dispose()
        }
        # the list of found files, written straight onto the picture
        if ($script:ListOverlay) {
            $o = $script:OverlayRect
            if ($o.Width -le 0) { return }
            $g.SmoothingMode = 'AntiAlias'; $g.TextRenderingHint = 'AntiAliasGridFit'
            $back = New-Object System.Drawing.SolidBrush ([System.Drawing.Color]::FromArgb(150,22,13,5)); $g.FillRectangle($back, $o); $back.Dispose()
            $edge = New-Object System.Drawing.Pen ([System.Drawing.Color]::FromArgb(200,196,150,72)), 1; $g.DrawRectangle($edge, $o); $edge.Dispose()
            $cols = @(@{ T='Файл'; X=0.0; W=0.42 }, @{ T='Где лежит'; X=0.42; W=0.38 }, @{ T='Тип'; X=0.80; W=0.06 }, @{ T='Изменён'; X=0.86; W=0.14 })
            $fmt = New-Object System.Drawing.StringFormat; $fmt.Trimming = 'EllipsisCharacter'; $fmt.FormatFlags = 'NoWrap'; $fmt.LineAlignment = 'Center'
            $goldB = New-Object System.Drawing.SolidBrush ([System.Drawing.Color]::FromArgb(247,214,140))
            $textB = New-Object System.Drawing.SolidBrush ([System.Drawing.Color]::FromArgb(255,246,226))
            $dimB = New-Object System.Drawing.SolidBrush ([System.Drawing.Color]::FromArgb(225,214,190))
            $iw = $o.Width - 24
            foreach ($c in $cols) { $g.DrawString($c.T, $script:ListHeadFont, $goldB, (New-Object System.Drawing.RectangleF ($o.X + 12 + $iw * $c.X), ($o.Y + 4), ($iw * $c.W - 8), 24), $fmt) }
            $line = New-Object System.Drawing.Pen ([System.Drawing.Color]::FromArgb(140,196,150,72)); $g.DrawLine($line, $o.X + 10, $o.Y + 30, $o.Right - 10, $o.Y + 30); $line.Dispose()
            $n = $grid.Items.Count; $vis = Get-OverlayVisibleRows
            $sig = [string]$n + '|' + $(if ($n) { [string]$grid.Items[0].Tag } else { '' })
            if ($sig -ne $script:OverlaySig) { $script:OverlaySig = $sig; $script:ListTopRow = 0 }   # a new search starts from the top
            if ($n -eq 0) { $g.DrawString('Пока пусто — найдите что-нибудь в строке сверху.', $script:ListFont, $dimB, (New-Object System.Drawing.RectangleF ($o.X + 14), ($o.Y + 36), ($iw), 26), $fmt) }
            $sel = if ($grid.SelectedIndices.Count) { $grid.SelectedIndices[0] } else { -1 }
            for ($k = 0; $k -lt $vis; $k++) {
                $i = $script:ListTopRow + $k; if ($i -ge $n) { break }
                $it = $grid.Items[$i]; $path = [string]$it.Tag
                $y = $o.Y + 32 + $k * $script:OverlayRowH
                if ($i -eq $sel) { $hl = New-Object System.Drawing.SolidBrush ([System.Drawing.Color]::FromArgb(150,196,150,72)); $g.FillRectangle($hl, $o.X + 4, $y, $o.Width - 8, $script:OverlayRowH); $hl.Dispose() }
                if (!$script:PlaceCache.ContainsKey($path)) { $script:PlaceCache[$path] = (Get-PlaceDescription $path) }
                $vals = @($it.Text, $script:PlaceCache[$path], $it.SubItems[2].Text, $it.SubItems[6].Text)
                for ($c = 0; $c -lt 4; $c++) {
                    $b = if ($c -eq 0) { $textB } else { $dimB }
                    $g.DrawString([string]$vals[$c], $script:ListFont, $b, (New-Object System.Drawing.RectangleF ($o.X + 12 + $iw * $cols[$c].X), $y, ($iw * $cols[$c].W - 8), $script:OverlayRowH), $fmt)
                }
            }
            if ($n -gt $vis) {
                $more = 'Показано ' + ($script:ListTopRow + 1) + '–' + [math]::Min($n, $script:ListTopRow + $vis) + ' из ' + $n + '   (колёсико мыши или ↑↓ — листать)'
                $g.DrawString($more, $script:ListFont, $goldB, (New-Object System.Drawing.RectangleF ($o.X + 12), ($o.Bottom - 24), ($iw), 22), $fmt)
            }
            $goldB.Dispose(); $textB.Dispose(); $dimB.Dispose(); $fmt.Dispose()
        }
    })
}

$script:AllFiles = @()
$script:CurrentRows = @()
function Show-Rows($Rows) {
    $script:CurrentRows = @($Rows)
    $grid.BeginUpdate()
    try {
        $grid.Items.Clear()
        foreach ($row in $script:CurrentRows) {
            $item = [System.Windows.Forms.ListViewItem]::new([string]$row.Name)
            [void]$item.SubItems.Add([string]$row.Category)
            [void]$item.SubItems.Add([string]$row.Extension)
            [void]$item.SubItems.Add([string]$row.SizeMB)
            [void]$item.SubItems.Add([string]$row.SizeBytes)
            [void]$item.SubItems.Add([string]$row.Created)
            [void]$item.SubItems.Add([string]$row.Modified)
            [void]$item.SubItems.Add([string]$row.ContentPreview)
            [void]$item.SubItems.Add([string]$row.ContentStatus)
            $item.Tag = [string]$row.Path
            [void]$grid.Items.Add($item)
        }
    } finally { $grid.EndUpdate() }
}
function Load-Index {
    if (Test-Path -LiteralPath $script:IndexPath) {
        try {
            $script:AllFiles = @(Import-Csv -LiteralPath $script:IndexPath -Encoding UTF8)
            # Older catalogs predate text indexing. Add an empty column in memory
            # so searching by name, extension, and path still works before indexing.
            foreach ($entry in $script:AllFiles) {
                if (!$entry.PSObject.Properties['ContentPreview']) { $entry | Add-Member -NotePropertyName ContentPreview -NotePropertyValue '' }
                if (!$entry.PSObject.Properties['ContentStatus']) { $entry | Add-Member -NotePropertyName ContentStatus -NotePropertyValue '' }
            }
        } catch { $script:AllFiles = @() }
    }
    Show-Rows $script:AllFiles
}
# Search understands an ordinary phrase («найди договор с Ивановым»): filler words are dropped,
# word endings are trimmed, words like «фото» or «таблица» narrow the file type.
$script:StopWords = @('найди','найти','найдите','поищи','поищите','поиск','покажи','покажите','открой','где','лежит','лежат','мне','нужно','нужен','нужна','надо','хочу','файл','файлы','файла','файлик','документ','документы','документа','пожалуйста','мой','моя','мои','мое','это','этот','эта','эту','эти','этой','можно','полностью','целиком','всю','весь','вообще','там','меня','был','была','было','были','который','которая','которое','про','о','об','по','с','со','и','в','во','на','за','для','из','от','к','а','у','же','ли','бы','тот','та','то','еще','какой','какая','привет')
$script:TypeWords = [ordered]@{ 'картин'='Изображения'; 'фото'='Изображения'; 'фотк'='Изображения'; 'скрин'='Изображения'; 'изображ'='Изображения'; 'видео'='Видео'; 'ролик'='Видео'; 'музык'='Аудио'; 'песн'='Аудио'; 'аудио'='Аудио'; 'таблиц'='Таблицы'; 'эксел'='Таблицы'; 'excel'='Таблицы'; 'презентац'='Презентации'; 'pdf'='PDF и книги'; 'пдф'='PDF и книги'; 'архив'='Архивы'; 'ворд'='Тексты'; 'word'='Тексты' }
function Get-SearchTerms([string]$Query) {
    $q = $Query.ToLowerInvariant().Replace('ё','е')
    $stems = New-Object System.Collections.Generic.List[string]
    $cats = New-Object System.Collections.Generic.List[string]
    $words = @([regex]::Split($q, '[^\p{L}\p{Nd}]+') | Where-Object { $_ })
    # «Библиотекарь, найди…» — an address, not part of the query.
    if ($words.Count -gt 1 -and $words[0].StartsWith('библиотекар')) { $words = @($words | Select-Object -Skip 1) }
    foreach ($w in $words) {
        if ($script:StopWords -contains $w) { continue }
        # Dictation fillers («э-э», «а», «мм», «ну», «вот») and stray single letters are not search words.
        if ($w -match '^(э+|а+|м+|у+|о+|ы+|хм+|эм+|ам+|ну|вот|типа|короче|значит|ой|ага|угу|так|ладно)$') { continue }
        if ($w.Length -lt 2 -and $w -notmatch '^\d$') { continue }
        $cat = $null
        foreach ($k in $script:TypeWords.Keys) { if ($w.StartsWith($k)) { $cat = $script:TypeWords[$k]; break } }
        if ($cat) { if (!$cats.Contains($cat)) { $cats.Add($cat) }; continue }
        $len = $w.Length
        $keep = if ($len -ge 6) { [math]::Max(4, $len - 3) } elseif ($len -eq 5) { 4 } else { $len }
        $stem = $w.Substring(0, $keep)
        if (!$stems.Contains($stem)) { $stems.Add($stem) }
    }
    return @{ Stems=@($stems); Categories=@($cats) }
}
function Get-PlaceDescription([string]$Path) {
    $dir = [IO.Path]::GetDirectoryName($Path)
    $nice = { param($s) $t = ([regex]::Replace([string]$s, '^\d+_', '')).Replace('_',' ').Trim().ToLowerInvariant(); if ($t) { $t.Substring(0,1).ToUpperInvariant() + $t.Substring(1) } else { $t } }
    $storage = Read-SmallTextFile $script:StorageRootConfigPath
    if ($storage -and $dir.StartsWith($storage.TrimEnd('\') + '\', [StringComparison]::OrdinalIgnoreCase)) {
        $parts = @($dir.Substring($storage.TrimEnd('\').Length + 1) -split '\\' | Where-Object { $_ })
        if ($parts.Count -eq 1) { return 'на полке «' + (& $nice $parts[0]) + '»' }
        return 'на полке «' + (& $nice $parts[0]) + '», в папке «' + (& $nice $parts[-1]) + '»'
    }
    $known = @(
        @{ P=(Get-DownloadsPath); T='в «Загрузках»' },
        @{ P=[Environment]::GetFolderPath('Desktop'); T='на Рабочем столе' },
        @{ P=[Environment]::GetFolderPath('MyDocuments'); T='в «Документах»' },
        @{ P=[Environment]::GetFolderPath('MyPictures'); T='в «Изображениях»' },
        @{ P=[Environment]::GetFolderPath('MyVideos'); T='в «Видео»' })
    foreach ($k in $known) {
        if ($k.P -and ($dir.TrimEnd('\') -eq $k.P.TrimEnd('\'))) { return $k.T }
        if ($k.P -and $dir.StartsWith($k.P.TrimEnd('\') + '\', [StringComparison]::OrdinalIgnoreCase)) { return $k.T + ', в папке «' + (& $nice ([IO.Path]::GetFileName($dir))) + '»' }
    }
    return 'в папке «' + (& $nice ([IO.Path]::GetFileName($dir))) + '»'
}
function Get-FilesWord([int]$N) {
    $m10 = $N % 10; $m100 = $N % 100
    if ($m10 -eq 1 -and $m100 -ne 11) { return 'файл' }
    if ($m10 -ge 2 -and $m10 -le 4 -and ($m100 -lt 12 -or $m100 -gt 14)) { return 'файла' }
    return 'файлов'
}
$script:HelperDir = Join-Path $script:MediaDir 'helper'
$script:HelperLengthMs = @{ hello=5100; search=5100; found=4600; dust=5100; secret=5100; bye=4300; sweep=6500; write=5100 }
$script:HelperTimer = New-Object System.Windows.Forms.Timer
$script:HelperTimer.Add_Tick({ $script:HelperTimer.Stop(); Show-HelperIdle })
function Set-HelperImage([string]$Path) {
    if (!(Test-Path -LiteralPath $Path)) { return }
    # Read into memory, so the picture file is never locked (it can be replaced while the program runs).
    try { $ms = New-Object System.IO.MemoryStream (,[IO.File]::ReadAllBytes($Path)); $old = $helperBox.Image; $helperBox.Image = [System.Drawing.Image]::FromStream($ms); if ($old) { $old.Dispose() } } catch { }
}
function Show-HelperIdle { $p = Join-Path $script:HelperDir 'idle_big.png'; if (!(Test-Path -LiteralPath $p)) { $p = Join-Path $script:HelperDir 'idle.png' }; Set-HelperImage $p }
function Show-Helper([string]$State, [switch]$Silent) {
    # hello, search (also «not found»), found, dust, secret, sweep, bye — a scene plays once, then he rests.
    Set-HelperImage (Join-Path $script:HelperDir ($State + '.gif'))
    $wav = Join-Path $script:HelperDir ($State + '.wav')
    if (!$Silent) {
        Play-Voice $wav
        # the live scene now follows this voice frame by frame
        if ($voiceCheck.Checked -and (Test-Path -LiteralPath $wav)) { $script:SceneUsesVoice = $true; if ($script:SceneClock) { $script:SceneClock.Restart() } }
    }
    $script:HelperTimer.Stop()
    $script:HelperTimer.Interval = [int]$(if ($script:HelperLengthMs.ContainsKey($State)) { $script:HelperLengthMs[$State] } else { 5000 })
    $script:HelperTimer.Start()
}
function Play-Voice([string]$Path) {
    # Arthur's lines go through the media player, not the system «PlaySound» channel:
    # a message-box beep used to cut him off mid-word.
    if (!$voiceCheck.Checked -or !(Test-Path -LiteralPath $Path)) { return }
    try {
        if (!$script:VoicePlayer) { Add-Type -AssemblyName PresentationCore; $script:VoicePlayer = New-Object System.Windows.Media.MediaPlayer }
        $script:VoicePlayer.Stop()
        $script:VoicePlayer.Open([Uri]$Path)
        $script:VoicePlayer.Volume = 1.0
        $script:VoicePlayer.Play()
    } catch {
        try { $sp = New-Object System.Media.SoundPlayer $Path; $sp.Play() } catch { }
    }
}
function Start-HelperLoop([string]$State, [switch]$Silent) {
    # For long jobs: the scene keeps repeating (with one line of voice at the start) until the job ends.
    Show-Helper $State -Silent:$Silent
    $script:HelperTimer.Stop()
}
function Play-Arthur([string]$Name, [switch]$Wait) {
    # Short phrases recorded in Arthur's voice (media\arthur_*.wav); the details are always written on screen.
    $path = Join-Path $script:MediaDir ('arthur_' + $Name + '.wav')
    if (!(Test-Path -LiteralPath $path)) { return }
    try {
        if ($script:ArthurPlayer) { $script:ArthurPlayer.Stop(); $script:ArthurPlayer.Dispose() }
        $script:ArthurPlayer = New-Object System.Media.SoundPlayer $path
        if ($Wait) { $script:ArthurPlayer.PlaySync() } else { $script:ArthurPlayer.Play() }
    } catch { }
}
function Apply-Search([switch]$Speak) {
    $q = $searchBox.Text.Trim()
    if (!$q) { Show-Rows $script:AllFiles; $status.Text = "В каталоге файлов: $($script:AllFiles.Count)"; $answerLabel.Text = 'Напишите, что найти, или нажмите «Сказать». Я отвечу здесь.'; return }
    $terms = Get-SearchTerms $q
    $stems = @($terms.Stems); $cats = @($terms.Categories)
    if ($stems.Count -eq 0 -and $cats.Count -eq 0) { $stems = @($q.ToLowerInvariant().Replace('ё','е')) }
    # A word must match from its beginning: «соба» finds «собака», but not «способа».
    $patterns = @($stems | ForEach-Object { New-Object System.Text.RegularExpressions.Regex ('(?<![\p{L}\p{Nd}])' + [regex]::Escape($_)), 'Compiled' })
    $scored = New-Object System.Collections.Generic.List[object]
    foreach ($e in $script:AllFiles) {
        $name = ([string]$e.Name).ToLowerInvariant().Replace('ё','е')
        $hay = ($name + ' ' + ([string]$e.Path) + ' ' + ([string]$e.ContentPreview)).ToLowerInvariant().Replace('ё','е')
        $hit = 0; $inName = 0
        foreach ($rx in $patterns) { if ($rx.IsMatch($hay)) { $hit++; if ($rx.IsMatch($name)) { $inName++ } } }
        if ($stems.Count -gt 0 -and $hit -eq 0) { continue }
        $scored.Add([pscustomobject]@{ E=$e; Hit=$hit; InName=$inName; CatOk=(($cats.Count -eq 0) -or ($cats -contains [string]$e.Category)); Time=[string]$e.Modified })
    }
    # All words first; if nothing has all of them, the files with most of the words.
    $found = @($scored | Where-Object { $_.Hit -eq $stems.Count })
    if ($found.Count -eq 0 -and $stems.Count -gt 1) { $found = @($scored | Where-Object { $_.Hit -ge [math]::Ceiling($stems.Count / 2) }) }
    if ($cats.Count -gt 0) { $withCat = @($found | Where-Object { $_.CatOk }); if ($withCat.Count -gt 0) { $found = $withCat } }
    $sorted = @($found | Sort-Object @{ Expression={ $_.Hit }; Descending=$true }, @{ Expression={ $_.InName }; Descending=$true }, @{ Expression={ $_.Time }; Descending=$true } | ForEach-Object { $_.E })
    Show-Rows $sorted
    $n = $sorted.Count
    if ($n -gt 0) {
        $grid.Items[0].Selected = $true; $grid.Items[0].Focused = $true; $grid.EnsureVisible(0)
        $place = Get-PlaceDescription ([string]$sorted[0].Path)
        $status.Text = "Найдено: $n. Первый: $($sorted[0].Name)"
        $answerLabel.Text = "Нашёл $n $(Get-FilesWord $n). " + $(if ($n -gt 1) { 'Первый (выделен) — ' } else { '' }) + '«' + [string]$sorted[0].Name + '» — лежит ' + $place + '. «Показать в папке» откроет это место, «Список файлов» покажет все найденные.'
        $sound = 'found'
    } else {
        $status.Text = "Ничего не найдено по запросу «$q»."
        $answerLabel.Text = "Ничего не нашёл по запросу «$q». Попробуйте другими словами или нажмите «Обновить каталог» — вдруг файл новый."
        # Looks like an instruction typed in search mode: point to the switch.
        $ql = $q.ToLowerInvariant()
        if ((Test-LooksLikeCommand $q) -or $ql -match '(созда|сдела|завед)\S*\s+(.*\s)?папк|уборк|собер\S*\s+.*(промпт|промт|канал)|открой\s+сайт') {
            $answerLabel.Text = "Похоже, это задание, а сейчас включён «Поиск». Нажмите кнопку «Задание» вверху и скажите это ещё раз — я выполню."
        }
        $sound = 'notfound'
    }
    # Found: «Ага! Вот ты где!» from the ladder scene; not found: «Так-так… Где же она у нас стояла?»
    if ($Speak -or $script:SpeakNext) { Show-Helper $(if ($sound -eq 'found') { 'found' } else { 'search' }) }
    $script:SpeakNext = $false
}
function Get-SelectedPaths {
    @($grid.SelectedItems | ForEach-Object { [string]$_.Tag } | Where-Object { $_ })
}
function Get-TypeFolder([string]$Extension) {
    switch (Get-Category $Extension) {
        'Изображения' { 'ИЗОБРАЖЕНИЯ' }
        'Видео' { 'ВИДЕО' }
        'Аудио' { 'МУЗЫКА' }
        'Тексты' { 'ТЕКСТЫ' }
        'Таблицы' { 'ТАБЛИЦЫ' }
        'Презентации' { 'ПРЕЗЕНТАЦИИ' }
        'PDF и книги' { 'PDF_И_КНИГИ' }
        'Архивы' { 'АРХИВЫ' }
        'Дизайн-проекты' { 'ДИЗАЙН_ПРОЕКТЫ' }
        'Код и данные' { 'КОД_И_ДАННЫЕ' }
        default { 'ДРУГИЕ_ФАЙЛЫ' }
    }
}
function Get-OrganizationProposal($Entry) {
    $path = [string]$Entry.Path
    $typeFolder = Get-TypeFolder ([string]$Entry.Extension)
    $segments = @($path -split '[\\/]+' | Where-Object { $_ })
    $lowerPath = $path.ToLowerInvariant()
    $destination = ''
    $reason = ''
    $confidence = 'Проверить вручную'
    # Own projects and rules from «настройки.json».
    $project = @($script:Projects | Where-Object { $lowerPath -match [string]$_.'Слова' }) | Select-Object -First 1
    $ownRule = @($script:ExtraRules | Where-Object { $lowerPath -match [string]$_.'Слова' }) | Select-Object -First 1

    if ($lowerPath -match 'парол|password|passport|паспорт|(^|[^а-яё-])банк|снилс|(^|[^а-яё])инн([^а-яё]|$)') {
        $destination = '04_ЛИЧНОЕ_И_СЕМЬЯ/ПРОВЕРИТЬ_ПРИВАТНОСТЬ/' + $typeFolder
        $reason = 'В названии или пути есть признак личных/чувствительных данных; перед переносом нужна ручная проверка.'
    } elseif ($project) {
        $title = if ($project.'Название') { [string]$project.'Название' } else { [string]$project.'Папка' }
        $destination = '02_МОИ_ПРОЕКТЫ/' + $project.'Папка' + '/' + $typeFolder
        $reason = 'Путь содержит название проекта «' + $title + '».'; $confidence = 'Хорошая подсказка, проверить'
    } elseif ($ownRule) {
        $destination = ([string]$ownRule.'Полка').Replace('\', '/').Trim('/') + '/' + $typeFolder
        $reason = 'Сработало ваше правило «' + $ownRule.'Слова' + '» из настроек.'; $confidence = 'Хорошая подсказка, проверить'
    } elseif ($lowerPath -match 'клиент|clients') {
        $clientName = ''
        for ($i = 0; $i -lt ($segments.Count - 1); $i++) {
            if ($segments[$i] -match '^(клиенты|клиент|clients)$') { $clientName = [string]$segments[$i + 1]; break }
        }
        if ($clientName) {
            $destination = '01_РАБОТА/КЛИЕНТЫ/' + $clientName + '/' + $typeFolder
            $reason = 'Путь содержит папку клиента «' + $clientName + '»; тему проекта стоит сверить.'; $confidence = 'Хорошая подсказка, проверить'
        } else {
            $destination = '90_РАЗОБРАТЬ/' + $typeFolder
            $reason = 'Есть слово «клиент», но имя папки клиента не удалось определить.'
        }
    } elseif ($lowerPath -match 'копирайт|тексты.?для.?клиент') {
        $destination = '01_РАБОТА/КОПИРАЙТИНГ/' + $typeFolder
        $reason = 'Путь содержит название направления «Копирайтинг».'; $confidence = 'Хорошая подсказка, проверить'
    } elseif ($lowerPath -match '(^|[^a-z])hr([^a-z]|$)|кадр|подбор') {
        $destination = '01_РАБОТА/HR/' + $typeFolder
        $reason = 'Путь содержит признаки HR/рабочих материалов.'; $confidence = 'Хорошая подсказка, проверить'
    } elseif ($lowerPath -match 'обучен|курс|урок|видеокурс|вайкодинг') {
        $destination = '05_ОБУЧЕНИЕ/' + $typeFolder
        $reason = 'Путь содержит признаки учебных материалов.'; $confidence = 'Хорошая подсказка, проверить'
    } elseif ($lowerPath -match 'маме|семь[яиеюё]|семей|(^|[^а-яё])личн|домашн|собак|кошк') {
        $destination = '04_ЛИЧНОЕ_И_СЕМЬЯ/' + $typeFolder
        $reason = 'Путь содержит признаки личных или семейных материалов.'; $confidence = 'Хорошая подсказка, проверить'
    } elseif ($lowerPath -match 'мои работы|портфолио|готовые работы') {
        $workName = ''
        for ($i = 0; $i -lt ($segments.Count - 1); $i++) {
            if ($segments[$i] -match '^(мои работы|портфолио|готовые работы)$') { $workName = [string]$segments[$i + 1]; break }
        }
        if ($workName) {
            $destination = '03_МОИ_РАБОТЫ/' + $workName + '/' + $typeFolder
            $reason = 'Путь содержит папку вашей работы «' + $workName + '».'; $confidence = 'Хорошая подсказка, проверить'
        } else {
            $destination = '03_МОИ_РАБОТЫ/' + $typeFolder
            $reason = 'Путь указывает на собственные работы или портфолио.'; $confidence = 'Проверить вручную'
        }
    } else {
        $destination = '90_РАЗОБРАТЬ/' + $typeFolder
        $reason = 'По доступным имени, типу и пути назначение неясно; отложить в папку разбора.'
    }

    # Only confident suggestions are pre-approved; everything else waits for the user's «да».
    $approve = if ($confidence -eq 'Хорошая подсказка, проверить') { 'да' } else { 'нет' }
    [pscustomobject]@{ 'Перенести'=$approve; 'ИсходныйПуть'=$path; 'Имя'=[string]$Entry.Name; 'Тип'=[string]$Entry.Extension; 'Размер'=(Format-Bytes ([long]$Entry.SizeBytes)); 'Изменён'=[string]$Entry.Modified; 'ПредлагаемаяПапка'=($script:LibraryRootName + '/' + $destination); 'Причина'=$reason; 'Уверенность'=$confidence }
}
# Folders on the scanned root (e.g. Desktop) move as a whole, keeping their inner structure.
# Pattern is matched against the whole top-level folder name (lower case, single spaces).
# Keep = keep the folder's own name inside the area; Sure = pre-approve with «да».
$script:FolderRules = @(
    [pscustomobject]@{ Pattern='^(клиенты|клиент|clients)$'; Area='01_РАБОТА/КЛИЕНТЫ'; Keep=$false; Sure=$true },
    [pscustomobject]@{ Pattern='^копирайтинг$'; Area='01_РАБОТА/КОПИРАЙТИНГ'; Keep=$false; Sure=$true },
    [pscustomobject]@{ Pattern='^(мои работы|портфолио|готовые работы)$'; Area='03_МОИ_РАБОТЫ'; Keep=$false; Sure=$true },
    [pscustomobject]@{ Pattern='^(личное|семья)$'; Area='04_ЛИЧНОЕ_И_СЕМЬЯ'; Keep=$false; Sure=$true },
    [pscustomobject]@{ Pattern='^маме$'; Area='04_ЛИЧНОЕ_И_СЕМЬЯ'; Keep=$true; Sure=$true },
    [pscustomobject]@{ Pattern='^(hr|кадры|подбор)$'; Area='01_РАБОТА/HR'; Keep=$false; Sure=$true },
    [pscustomobject]@{ Pattern='^(обучение|курсы)$'; Area='05_ОБУЧЕНИЕ'; Keep=$false; Sure=$true },
    # Guesses by name only: suggested, but left at «нет» until the user agrees.
    [pscustomobject]@{ Pattern='ваканси|резюме|отклик|самозанят|выплат|расч[её]т|фриланс'; Area='01_РАБОТА'; Keep=$true; Sure=$false },
    [pscustomobject]@{ Pattern='психолог|обучен|урок|вайкодинг|^курс'; Area='05_ОБУЧЕНИЕ'; Keep=$true; Sure=$false },
    [pscustomobject]@{ Pattern='конкурс|кейс|мультфильм'; Area='03_МОИ_РАБОТЫ'; Keep=$true; Sure=$false }
)
# Own projects and rules from «настройки.json» are checked first.
$ownFolderRules = @($script:Projects | ForEach-Object { [pscustomobject]@{ Pattern=[string]$_.'Слова'; Area='02_МОИ_ПРОЕКТЫ'; Keep=$true; Sure=$true } }) +
    @($script:ExtraRules | ForEach-Object { [pscustomobject]@{ Pattern=[string]$_.'Слова'; Area=([string]$_.'Полка').Replace('\', '/').Trim('/'); Keep=$true; Sure=$true } })
$script:FolderRules = @($ownFolderRules) + @($script:FolderRules)
function Get-PlanRow($Entry, $RootPrefixes) {
    $path = [string]$Entry.Path
    $base = ''
    foreach ($prefix in $RootPrefixes) { if ($path.StartsWith($prefix, [StringComparison]::OrdinalIgnoreCase)) { $base = $prefix; break } }
    $relative = if ($base) { $path.Substring($base.Length) } else { [IO.Path]::GetFileName($path) }
    $parts = @($relative -split '\\' | Where-Object { $_ })
    # @() around the whole if: otherwise a single folder collapses into a string and $dirs[0] is its first letter.
    $dirs = @(if ($parts.Count -gt 1) { $parts[0..($parts.Count - 2)] })

    if ($dirs.Count -eq 0) {
        # A loose file: sort by topic and type, as before.
        $row = Get-OrganizationProposal $Entry
        $group = '(файлы без папки)'; $groupTarget = '(по типам файлов)'
        $lowerName = ([string]$Entry.Name).ToLowerInvariant()
        if ($lowerName -eq 'desktop.ini' -or ([string]$Entry.Extension) -in @('.lnk','.url')) {
            $row.'Перенести' = 'нет'; $row.'Уверенность' = 'Проверить вручную'
            $row.'Причина' = 'Ярлык или служебный файл Windows — обычно его оставляют на месте.'
        }
    } else {
        $top = [string]$dirs[0]
        $rest = @($dirs | Select-Object -Skip 1)
        $name = ([regex]::Replace($top.ToLowerInvariant(), '\s+', ' ')).Trim()
        $rule = $script:FolderRules | Where-Object { $name -match $_.Pattern } | Select-Object -First 1
        if ($rule) {
            $groupParts = @($rule.Area) + $(if ($rule.Keep) { @($top) } else { @() })
            $sure = [bool]$rule.Sure
            $reason = if ($sure) { 'Папка «' + $top + '» переезжает целиком; структура внутри сохраняется.' } else { 'Папка «' + $top + '»: раздел угадан по названию — проверьте. Структура внутри сохраняется.' }
        } else {
            $groupParts = @('90_РАЗОБРАТЬ', $top)
            $sure = $false
            $reason = 'Назначение папки «' + $top + '» неясно; она целиком предложена в папку разбора.'
        }
        $groupTarget = $script:LibraryRootName + '/' + ($groupParts -join '/')
        $destination = (@($groupParts) + @($rest)) -join '/'
        $confidence = if ($sure) { 'Хорошая подсказка, проверить' } else { 'Проверить вручную' }
        $approve = if ($sure) { 'да' } else { 'нет' }
        if ($relative.ToLowerInvariant() -match 'парол|password|passport|паспорт|(^|[^а-яё-])банк|снилс|(^|[^а-яё])инн([^а-яё]|$)') {
            $approve = 'нет'; $confidence = 'Проверить вручную'
            $reason = 'Похоже на личные данные (пароли, документы) — файл останется на месте, пока вы сами не поставите «да». ' + $reason
        }
        $group = $top
        $row = [pscustomobject]@{ 'Перенести'=$approve; 'ИсходныйПуть'=$path; 'Имя'=[string]$Entry.Name; 'Тип'=[string]$Entry.Extension; 'Размер'=(Format-Bytes ([long]$Entry.SizeBytes)); 'Изменён'=[string]$Entry.Modified; 'ПредлагаемаяПапка'=($script:LibraryRootName + '/' + $destination); 'Причина'=$reason; 'Уверенность'=$confidence }
    }
    $row | Add-Member -NotePropertyName 'Группа' -NotePropertyValue $group
    $row | Add-Member -NotePropertyName 'КудаГруппа' -NotePropertyValue $groupTarget
    return $row
}
function Write-OrganizationPlan($Paths) {
    $byPath = @{}
    foreach ($entry in $script:AllFiles) { $byPath[[string]$entry.Path] = $entry }
    # «~$имя» — служебные файлы-замки открытых документов Office; переносить их нельзя.
    $selected = @($Paths | ForEach-Object { $byPath[[string]$_] } | Where-Object { $_ -and ([string]$_.Name) -notlike '~$*' })
    if ($selected.Count -eq 0) { [void][LibrarianApp.LibBox]::Show('Не нашла выбранные файлы в каталоге. Нажми «Обновить каталог» и попробуй снова.'); return }
    # Longest root first, so a nested catalog folder wins over its parent.
    $rootPrefixes = @(@($rootsBox.Items) | ForEach-Object { ([string]$_).TrimEnd('\') + '\' } | Sort-Object Length -Descending)
    $rows = @($selected | ForEach-Object { Get-PlanRow $_ $rootPrefixes })
    $csvPath = $script:OrganizationPlanCsvPath
    $htmlPath = Join-Path $script:DataHome 'organization_plan.html'
    # Keep the previous (possibly hand-edited) plan instead of silently overwriting it.
    if (Test-Path -LiteralPath $csvPath) { Copy-Item -LiteralPath $csvPath -Destination (Join-Path $script:DataHome 'organization_plan_previous.csv') -Force }
    $rows | Select-Object 'Перенести','Группа','ИсходныйПуть','Имя','Тип','Размер','Изменён','ПредлагаемаяПапка','Причина','Уверенность' |
        Export-Csv -LiteralPath $csvPath -NoTypeInformation -Encoding UTF8 -Delimiter ';'
    $enc = { param($v) [System.Net.WebUtility]::HtmlEncode([string]$v) }

    $groupRows = New-Object System.Text.StringBuilder
    foreach ($g in ($rows | Group-Object 'Группа' | Sort-Object Count -Descending)) {
        $yes = @($g.Group | Where-Object { $_.'Перенести' -eq 'да' }).Count
        $cls = if ($yes -eq $g.Count) { 'yes' } elseif ($yes -eq 0) { 'no' } else { 'part' }
        [void]$groupRows.Append('<tr class="' + $cls + '"><td>' + (& $enc $g.Name) + '</td><td>' + $g.Count + '</td><td>' + $yes + '</td><td>' + (& $enc $g.Group[0].'КудаГруппа') + '</td><td>' + (& $enc $g.Group[0].'Причина') + '</td></tr>')
    }
    $tableRows = New-Object System.Text.StringBuilder
    foreach ($row in $rows) {
        $cells = @($row.'Перенести',$row.'Группа',$row.'ИсходныйПуть',$row.'Тип',$row.'Размер',$row.'ПредлагаемаяПапка',$row.'Причина')
        [void]$tableRows.Append('<tr>')
        foreach ($cell in $cells) { [void]$tableRows.Append('<td>' + (& $enc $cell) + '</td>') }
        [void]$tableRows.Append('</tr>')
    }
    $approvedCount = @($rows | Where-Object { $_.'Перенести' -eq 'да' }).Count
    $html = @"
<!doctype html><html lang="ru"><meta charset="utf-8"><title>План раскладки файлов</title>
<style>body{font:15px Segoe UI,Arial,sans-serif;margin:28px;color:#292521;background:#faf8f3}h1,h2{color:#4c3b25}p{line-height:1.5}.note{padding:12px;background:#fff1d6;border-left:5px solid #d49b36}table{border-collapse:collapse;width:100%;background:white;margin-bottom:28px}th,td{border:1px solid #d8d2c8;padding:7px;text-align:left;vertical-align:top}th{background:#eee7db;position:sticky;top:0}td{overflow-wrap:anywhere}tr.yes td:first-child{border-left:5px solid #5a9a4c}tr.no td:first-child{border-left:5px solid #c9c2b6}tr.part td:first-child{border-left:5px solid #d49b36}</style>
<h1>План раскладки файлов</h1>
<p>Файлов в плане: $($rows.Count), из них отмечено «да»: $approvedCount. Папки переезжают целиком — всё, что внутри, сохраняет свою структуру. По типам (ИЗОБРАЖЕНИЯ, ТЕКСТЫ…) раскладываются только файлы, лежащие без папки.</p>
<p class="note"><b>Это только план.</b> Ничего не перемещено. Зелёная полоса — папка отмечена «да» целиком, серая — «нет», жёлтая — частично (например, файлы с паролями или документами оставлены на месте).</p>
<p><b>Как выполнить:</b> нажмите в программе «Выполнить план…» и отметьте галочками папки, которые нужно перенести сейчас. Excel не обязателен — он нужен, только если хотите поправить отдельные файлы (кнопка «Открыть план в Excel»; сохранять тем же файлом, формат CSV). Последний перенос отменяется кнопкой «Отменить перенос».</p>
<h2>Сводка по папкам</h2>
<table><thead><tr><th>Папка</th><th>Файлов</th><th>Из них «да»</th><th>Куда</th><th>Почему</th></tr></thead><tbody>
$($groupRows.ToString())
</tbody></table>
<h2>Все файлы</h2>
<table><thead><tr><th>Перенести</th><th>Папка</th><th>Исходный путь</th><th>Тип</th><th>Размер</th><th>Куда</th><th>Почему</th></tr></thead><tbody>
$($tableRows.ToString())
</tbody></table></html>
"@
    Set-Content -LiteralPath $htmlPath -Value $html -Encoding UTF8
    $script:OrganizationPlanPath = $htmlPath
    $status.Text = "План готов: $($rows.Count) файлов, отмечено «да»: $approvedCount. Ничего не перемещено."
    if (!(Open-InBrowser $htmlPath)) { [void][LibrarianApp.LibBox]::Show("План сохранён, но браузер не открылся.`r`n$htmlPath`r`nТаблица CSV: $csvPath") }
}
function Open-InBrowser([string]$Path) {
    # .html may be associated with Notepad, so ask Windows for the default web browser instead.
    try {
        $progId = (Get-ItemProperty -LiteralPath 'HKCU:\Software\Microsoft\Windows\Shell\Associations\UrlAssociations\http\UserChoice' -ErrorAction Stop).ProgId
        $command = [string](Get-ItemProperty -LiteralPath ("Registry::HKEY_CLASSES_ROOT\$progId\shell\open\command") -ErrorAction Stop).'(default)'
        if ($command -match '^\s*"([^"]+)"\s*(.*)$' -or $command -match '^\s*(\S+)\s*(.*)$') {
            $exe = $Matches[1]; $arguments = $Matches[2]
            # A file:/// URI has no spaces, so it needs no quotes. Quotes break browsers registered
            # with «--single-argument %1» (Yandex, Chrome): they take the quotes as part of the address.
            $uri = ([Uri]$Path).AbsoluteUri
            $arguments = if ($arguments -match '%1') { $arguments.Replace('%1', $uri) } else { ($arguments + ' ' + $uri).Trim() }
            if (Test-Path -LiteralPath $exe -PathType Leaf) { Start-Process -FilePath $exe -ArgumentList $arguments -ErrorAction Stop; return $true }
        }
    } catch { }
    try { Start-Process -FilePath 'msedge.exe' -ArgumentList ('"' + $Path + '"') -ErrorAction Stop; return $true } catch { }
    try { Start-Process -FilePath $Path -ErrorAction Stop; return $true } catch { return $false }
}
function Open-PlanInExcel {
    $csv = $script:OrganizationPlanCsvPath
    if (!(Test-Path -LiteralPath $csv)) { [void][LibrarianApp.LibBox]::Show('Плана ещё нет. Сначала нажми «План раскладки».'); return }
    try { Start-Process -FilePath 'excel.exe' -ArgumentList ('"' + $csv + '"') -ErrorAction Stop }
    catch { [void][LibrarianApp.LibBox]::Show("Не удалось запустить Excel. Открой таблицу вручную (правой кнопкой → Открыть с помощью → Excel):`r`n$csv") }
}
function Read-PlanCsv([string]$Path) {
    # Excel re-saves CSV in the ANSI code page unless «CSV UTF-8» is chosen; accept both.
    $bytes = [System.IO.File]::ReadAllBytes($Path)
    if ($bytes.Length -ge 2 -and $bytes[0] -eq 0xFF -and $bytes[1] -eq 0xFE) { $text = [System.Text.Encoding]::Unicode.GetString($bytes, 2, $bytes.Length - 2) }
    else {
        try { $text = (New-Object System.Text.UTF8Encoding($false, $true)).GetString($bytes) }
        catch { $text = [System.Text.Encoding]::Default.GetString($bytes) }
    }
    $lines = @($text.TrimStart([char]0xFEFF) -split "`r?`n" | Where-Object { $_ -ne '' })
    if ($lines.Count -eq 0) { return @() }
    $delimiter = if ($lines[0] -match "`t") { "`t" } elseif ($lines[0] -notmatch ';' -and $lines[0] -match ',') { ',' } else { ';' }
    @($lines | ConvertFrom-Csv -Delimiter $delimiter)
}
function Get-PlanGroups($Rows) {
    @($Rows | Group-Object { if ($_.PSObject.Properties['Группа'] -and $_.'Группа') { [string]$_.'Группа' } else { '(без группы)' } } | Sort-Object Count -Descending | ForEach-Object {
        $shortest = $_.Group | Sort-Object { ([string]$_.'ПредлагаемаяПапка').Length } | Select-Object -First 1
        $target = if ($_.Name -eq '(файлы без папки)') { 'по типам файлов' } else { [string]$shortest.'ПредлагаемаяПапка' }
        [pscustomobject]@{ Name=$_.Name; Count=$_.Count; Target=$target; Rows=@($_.Group) }
    })
}
function Select-PlanGroups($Rows) {
    # Excel filters only hide rows, so a «да» in a hidden folder would still move.
    # The user explicitly ticks the folders to move now; nothing is ticked by default.
    $groups = @(Get-PlanGroups $Rows)
    $dlg = New-Object System.Windows.Forms.Form
    $dlg.Text = 'Какие папки перенести сейчас?'; $dlg.StartPosition = 'CenterParent'; $dlg.Font = $form.Font
    $dlg.Size = [System.Drawing.Size]::new(760,520); $dlg.FormBorderStyle = 'FixedDialog'; $dlg.MinimizeBox = $false; $dlg.MaximizeBox = $false
    $hint = New-Object System.Windows.Forms.Label
    $hint.Text = 'Отметь папки, которые нужно перенести сейчас. Для первого раза лучше одну небольшую. Файлы, у которых в плане «нет» (например, с паролями и документами), останутся на месте.'
    $hint.Location = [System.Drawing.Point]::new(12,10); $hint.Size = [System.Drawing.Size]::new(720,40)
    $dlg.Controls.Add($hint)
    $list = New-Object System.Windows.Forms.CheckedListBox
    $list.CheckOnClick = $true; $list.HorizontalScrollbar = $true
    $list.Location = [System.Drawing.Point]::new(12,55); $list.Size = [System.Drawing.Size]::new(720,370)
    foreach ($g in $groups) { [void]$list.Items.Add(('{0} — {1} файлов → {2}' -f $g.Name, $g.Count, $g.Target)) }
    $dlg.Controls.Add($list)
    $ok = New-Object System.Windows.Forms.Button
    $ok.Text = 'Перенести отмеченные'; $ok.Location = [System.Drawing.Point]::new(462,438); $ok.Size = [System.Drawing.Size]::new(170,30)
    $ok.DialogResult = [System.Windows.Forms.DialogResult]::OK
    $cancel = New-Object System.Windows.Forms.Button
    $cancel.Text = 'Отмена'; $cancel.Location = [System.Drawing.Point]::new(640,438); $cancel.Size = [System.Drawing.Size]::new(92,30)
    $cancel.DialogResult = [System.Windows.Forms.DialogResult]::Cancel
    $dlg.Controls.Add($ok); $dlg.Controls.Add($cancel)
    $dlg.AcceptButton = $ok; $dlg.CancelButton = $cancel
    try {
        if ($dlg.ShowDialog($form) -ne [System.Windows.Forms.DialogResult]::OK) { return @{ Ok=$false; Rows=@() } }
        $picked = New-Object System.Collections.Generic.List[object]
        foreach ($i in $list.CheckedIndices) { foreach ($r in $groups[$i].Rows) { $picked.Add($r) } }
        return @{ Ok=$true; Rows=$picked.ToArray() }
    } finally { $dlg.Dispose() }
}
function Get-SafeRelativeFolder([string]$Proposed) {
    # Accept only a relative path inside the storage root; reject drives, UNC paths and «..».
    $value = ([string]$Proposed).Trim()
    if (!$value -or $value -match ':' -or $value.StartsWith('\') -or $value.StartsWith('/')) { return $null }
    $parts = @($value -split '[\\/]+' | ForEach-Object { $_.Trim() } | Where-Object { $_ })
    if ($parts.Count -gt 0 -and (@($script:LibraryRootName) + $script:OldLibraryRootNames) -contains $parts[0]) { $parts = @($parts | Select-Object -Skip 1) }
    if ($parts.Count -eq 0) { return $null }
    $invalid = [IO.Path]::GetInvalidFileNameChars()
    $clean = New-Object System.Collections.Generic.List[string]
    foreach ($part in $parts) {
        if ($part -eq '.' -or $part -eq '..') { return $null }
        $name = (-join @($part.ToCharArray() | ForEach-Object { if ($invalid -contains $_) { '_' } else { [string]$_ } })).TrimEnd('.',' ')
        if (!$name) { return $null }
        $clean.Add($name)
    }
    return ($clean -join '\')
}
function Read-SmallTextFile([string]$Path) {
    # A power loss right after a write can leave a file filled with zero bytes; treat them as nothing.
    if (!(Test-Path -LiteralPath $Path -PathType Leaf)) { return '' }
    try { return ([IO.File]::ReadAllText($Path, [Text.Encoding]::UTF8)).Replace([string][char]0, '').Trim([char]0xFEFF, ' ', "`r", "`n", "`t") } catch { return '' }
}
function Get-StorageRoot {
    $saved = ''
    if (Test-Path -LiteralPath $script:StorageRootConfigPath) { $saved = Read-SmallTextFile $script:StorageRootConfigPath }
    $dialog = New-Object System.Windows.Forms.FolderBrowserDialog
    $dialog.Description = "Выбери, ГДЕ должны стоять книжные полки (например, Документы или диск D), или саму папку $($script:LibraryRootName), если она уже есть. Папку $($script:LibraryRootName) программа создаст сама."
    # A folder created inside this dialog cannot be renamed while the dialog is open
    # («Папка уже используется»), so the program creates the storage folder itself.
    $dialog.ShowNewFolderButton = $false
    # A folder saved under an old root name starts the dialog one level up.
    if ($saved -and $script:OldLibraryRootNames -contains [IO.Path]::GetFileName($saved.TrimEnd('\'))) { $saved = [IO.Path]::GetDirectoryName($saved.TrimEnd('\')) }
    if ($saved -and (Test-Path -LiteralPath $saved -PathType Container)) { $dialog.SelectedPath = $saved }
    if ($dialog.ShowDialog() -ne 'OK') { return $null }
    $resolved = Resolve-StorageRoot $dialog.SelectedPath
    if ($resolved) { Add-StorageRootToCatalog }
    return $resolved
}
function Add-StorageRootToCatalog {
    # Without this, «Обновить каталог» would drop moved files from search.
    if (!(Test-Path -LiteralPath $script:StorageRootConfigPath)) { return }
    $storage = Read-SmallTextFile $script:StorageRootConfigPath
    if (!$storage -or !(Test-Path -LiteralPath $storage -PathType Container)) { return }
    $prefix = $storage.TrimEnd('\') + '\'
    foreach ($existing in @($rootsBox.Items)) {
        $existingPrefix = ([string]$existing).TrimEnd('\') + '\'
        if ($prefix.StartsWith($existingPrefix, [StringComparison]::OrdinalIgnoreCase)) { return }
    }
    [void]$rootsBox.Items.Add($storage)
    Save-Roots @($rootsBox.Items)
}
function Resolve-StorageRoot([string]$Selected) {
    $root = $Selected
    # A folder inside existing shelves (e.g. КНИЖНЫЕ_ПОЛКИ\04_ЛИЧНОЕ_И_СЕМЬЯ) means those shelves;
    # otherwise a second КНИЖНЫЕ_ПОЛКИ would be created inside the first one.
    $probe = $Selected.TrimEnd('\')
    while ($probe) {
        if ([IO.Path]::GetFileName($probe) -eq $script:LibraryRootName) {
            Set-Content -LiteralPath $script:StorageRootConfigPath -Value $probe -Encoding UTF8
            Add-EmptyShelves $probe
            return $probe
        }
        $parent = [IO.Path]::GetDirectoryName($probe)
        if (!$parent -or $parent -eq $probe) { break }
        $probe = $parent
    }
    if ([IO.Path]::GetFileName($root.TrimEnd('\')) -ne $script:LibraryRootName) {
        $candidate = Join-Path $root $script:LibraryRootName
        $exists = if (Test-Path -LiteralPath $candidate -PathType Container) { 'уже есть' } else { 'будет создана' }
        $answer = [LibrarianApp.LibBox]::Show("Книжные полки: $candidate`r`n(папка $exists)`r`n`r`nДа — расставить группы (01_РАБОТА, 02_МОИ_ПРОЕКТЫ…) в папке $($script:LibraryRootName).`r`nНет — расставить группы прямо в выбранной папке:`r`n$root", 'Где стоят книжные полки', 'YesNoCancel', 'Question')
        if ($answer -eq [Windows.Forms.DialogResult]::Cancel) { return $null }
        if ($answer -eq [Windows.Forms.DialogResult]::Yes) {
            try { [void][IO.Directory]::CreateDirectory($candidate) } catch { [void][LibrarianApp.LibBox]::Show("Не удалось создать папку:`r`n$candidate`r`n$($_.Exception.Message)"); return $null }
            $root = $candidate
        }
    }
    Set-Content -LiteralPath $script:StorageRootConfigPath -Value $root -Encoding UTF8
    Add-EmptyShelves $root
    return $root
}
function Add-EmptyShelves([string]$Root) {
    # Empty shelves right away: «Новая папка» and the sorting plan offer only shelves that exist.
    foreach ($shelf in @('01_РАБОТА', '02_МОИ_ПРОЕКТЫ', '03_МОИ_РАБОТЫ', '04_ЛИЧНОЕ_И_СЕМЬЯ', '05_ОБУЧЕНИЕ', '90_РАЗОБРАТЬ')) {
        try { [void][IO.Directory]::CreateDirectory((Join-Path $Root $shelf)) } catch { }
    }
}
function Add-MoveLogEntry([string]$LogPath, [string]$State, [string]$Source, [string]$Target, [string]$Note) {
    # Appended row by row and written through the disk cache, so the undo log survives
    # even a sudden reboot in the middle of a move (otherwise the file can end up as zero bytes).
    $isNew = !(Test-Path -LiteralPath $LogPath)
    $lines = @([pscustomobject]@{ Status=$State; Source=$Source; Target=$Target; Note=$Note; Time=(Get-Date -Format 'yyyy-MM-dd HH:mm:ss') } | ConvertTo-Csv -NoTypeInformation -Delimiter ';')
    if (!$isNew) { $lines = @($lines | Select-Object -Skip 1) }
    $stream = New-Object System.IO.FileStream($LogPath, [System.IO.FileMode]::Append, [System.IO.FileAccess]::Write, [System.IO.FileShare]::Read, 4096, [System.IO.FileOptions]::WriteThrough)
    try {
        $writer = New-Object System.IO.StreamWriter($stream, (New-Object System.Text.UTF8Encoding($true)))
        foreach ($line in $lines) { $writer.WriteLine($line) }
        $writer.Flush(); $stream.Flush($true)
    } finally { $stream.Dispose() }
}
function Read-MoveLog([string]$Path) {
    # Tolerates zero bytes left by a power loss: damaged lines are skipped, the rest is still usable.
    $text = ([IO.File]::ReadAllText($Path, [Text.Encoding]::UTF8)).Replace([string][char]0, '')
    $lines = @($text.TrimStart([char]0xFEFF) -split "`r?`n" | Where-Object { $_.Trim() })
    if ($lines.Count -eq 0) { return @() }
    @($lines | ConvertFrom-Csv -Delimiter ';' | Where-Object { $_.PSObject.Properties['Status'] -and $_.Status })
}
function Get-MoveTasks($Rows, [string]$StorageRoot) {
    $tasks = New-Object System.Collections.Generic.List[object]
    $problems = New-Object System.Collections.Generic.List[object]
    foreach ($row in $Rows) {
        $source = ([string]$row.'ИсходныйПуть').Trim()
        $relative = Get-SafeRelativeFolder ([string]$row.'ПредлагаемаяПапка')
        if (!$relative) { $problems.Add([pscustomobject]@{ Source=$source; Note=('Непонятная папка назначения: ' + $row.'ПредлагаемаяПапка') }); continue }
        if (!(Test-Path -LiteralPath $source -PathType Leaf)) { $problems.Add([pscustomobject]@{ Source=$source; Note='Файл не найден на прежнем месте.' }); continue }
        $tasks.Add([pscustomobject]@{ Source=$source; Folder=(Join-Path $StorageRoot $relative) })
    }
    return @{ Tasks=$tasks; Problems=$problems }
}
function Invoke-MoveTasks($Tasks, [string]$LogPath) {
    # Moves without overwriting (a number is added on a name clash); every step goes to the undo log.
    $byPath = @{}
    foreach ($entry in $script:AllFiles) { $byPath[[string]$entry.Path] = $entry }
    $moved = 0; $skipped = 0; $failed = 0; $i = 0
    $errors = New-Object System.Collections.Generic.List[string]
    foreach ($task in $Tasks) {
        $i++
        try {
            $sourceDir = [IO.Path]::GetDirectoryName($task.Source)
            if ([string]::Equals($sourceDir.TrimEnd('\'), $task.Folder.TrimEnd('\'), [StringComparison]::OrdinalIgnoreCase)) {
                $skipped++; Add-MoveLogEntry $LogPath 'ПРОПУЩЕН' $task.Source '' 'Файл уже лежит в этой папке.'; continue
            }
            [void][IO.Directory]::CreateDirectory($task.Folder)
            $fileName = [IO.Path]::GetFileName($task.Source)
            $target = Join-Path $task.Folder $fileName
            if (Test-Path -LiteralPath $target) {
                $base = [IO.Path]::GetFileNameWithoutExtension($fileName); $ext = [IO.Path]::GetExtension($fileName); $n = 1
                do { $target = Join-Path $task.Folder ("{0} ({1}){2}" -f $base,$n,$ext); $n++ } while (Test-Path -LiteralPath $target)
            }
            Move-Item -LiteralPath $task.Source -Destination $target -ErrorAction Stop
            $moved++
            Add-MoveLogEntry $LogPath 'ПЕРЕНЕСЁН' $task.Source $target ''
            $entry = $byPath[$task.Source]
            if ($entry) { $entry.Path = $target; $entry.Name = [IO.Path]::GetFileName($target) }
        } catch {
            $failed++; $errors.Add($_.Exception.Message); Add-MoveLogEntry $LogPath 'ОШИБКА' $task.Source '' $_.Exception.Message
        }
        if (($i % 10) -eq 0 -or $i -eq $Tasks.Count) { $status.Text = "Переношу: $i из $($Tasks.Count)…"; $form.Refresh(); [Windows.Forms.Application]::DoEvents() }
    }
    return [pscustomobject]@{ Moved=$moved; Skipped=$skipped; Failed=$failed; Errors=$errors }
}
function Invoke-OrganizationPlan {
    $csvPath = $script:OrganizationPlanCsvPath
    if (!(Test-Path -LiteralPath $csvPath)) { [void][LibrarianApp.LibBox]::Show('Плана ещё нет. Сначала нажми «План раскладки».'); return }
    try { $rows = @(Read-PlanCsv $csvPath) } catch { [void][LibrarianApp.LibBox]::Show("Не удалось прочитать план:`r`n$csvPath`r`n`r`nПричина: $($_.Exception.Message)`r`nЕсли файл открыт в Excel, закрой его и попробуй снова."); return }
    if ($rows.Count -eq 0 -or !$rows[0].PSObject.Properties['Перенести'] -or !$rows[0].PSObject.Properties['ИсходныйПуть'] -or !$rows[0].PSObject.Properties['ПредлагаемаяПапка']) {
        [void][LibrarianApp.LibBox]::Show('План в старом формате или повреждён. Создай его заново кнопкой «План раскладки».'); return
    }
    $approved = @($rows | Where-Object { ([string]$_.'Перенести').Trim().ToLowerInvariant() -in @('да','д','yes','y','+','1') })
    if ($approved.Count -eq 0) { [void][LibrarianApp.LibBox]::Show("В плане нет строк с «да» в колонке «Перенести».`r`nНажми «Открыть план в Excel», поставь «да» нужным строкам, сохрани и закрой Excel."); return }
    $choice = Select-PlanGroups $approved
    if (!$choice.Ok) { return }
    $approved = @($choice.Rows)
    if ($approved.Count -eq 0) { [void][LibrarianApp.LibBox]::Show('Не отмечено ни одной папки — ничего не перенесено.'); return }

    $storageRoot = Get-StorageRoot
    if (!$storageRoot) { return }
    $prepared = Get-MoveTasks $approved $storageRoot
    $tasks = $prepared.Tasks; $problems = $prepared.Problems
    if ($tasks.Count -eq 0) { [void][LibrarianApp.LibBox]::Show("Нечего переносить: у всех отмеченных строк нет файла или папка назначения указана неверно ($($problems.Count)). Обнови каталог и пересоздай план."); return }
    $folderLines = (@(Get-PlanGroups $approved | Select-Object -First 12 | ForEach-Object { '  • ' + $_.Name + ' — ' + $_.Count }) -join "`r`n")
    $message = "Будет перенесено файлов: $($tasks.Count)`r`nПапки:`r`n$folderLines`r`nКуда: $storageRoot`r`nБудет пропущено: $($problems.Count)`r`n`r`nФайлы не перезаписываются: при совпадении имён к новому имени добавится номер. Каждый перенос записывается в журнал, и его можно отменить кнопкой «Отменить перенос».`r`n`r`nПеренести?"
    if ([LibrarianApp.LibBox]::Show($message, 'Подтверждение переноса', 'YesNo', 'Warning') -ne 'Yes') { return }

    [void][IO.Directory]::CreateDirectory($script:MoveLogDir)
    $logPath = Join-Path $script:MoveLogDir ('moves_' + (Get-Date -Format 'yyyyMMdd_HHmmss') + '.csv')
    foreach ($problem in $problems) { Add-MoveLogEntry $logPath 'ПРОПУЩЕН' $problem.Source '' $problem.Note }
    $organizeButton.Enabled = $false; $undoMoveButton.Enabled = $false; $scanButton.Enabled = $false
    try {
        $result = Invoke-MoveTasks $tasks $logPath
        $moved = $result.Moved; $skipped = $problems.Count + $result.Skipped; $failed = $result.Failed
    } finally {
        $organizeButton.Enabled = $true; $undoMoveButton.Enabled = $true; $scanButton.Enabled = $true
        $script:AllFiles | Export-Csv -LiteralPath $script:IndexPath -NoTypeInformation -Encoding UTF8
        Reset-DuplicateScan
        Apply-Search
    }
    $status.Text = "Перенесено: $moved; пропущено: $skipped; ошибок: $failed."
    [void][LibrarianApp.LibBox]::Show("Перенесено: $moved`r`nПропущено: $skipped`r`nОшибок: $failed`r`n`r`nЖурнал: $logPath`r`nВернуть всё обратно можно кнопкой «Отменить перенос».", 'Перенос завершён')
}
function Undo-LastMove {
    $logs = @(Get-ChildItem -LiteralPath $script:MoveLogDir -Filter 'moves_*.csv' -File -ErrorAction SilentlyContinue | Where-Object { $_.Name -notlike '*.undone.csv' } | Sort-Object Name -Descending)
    if ($logs.Count -eq 0) { [void][LibrarianApp.LibBox]::Show('Нет переносов, которые можно отменить.'); return }
    $log = $logs[0]
    try { $entries = @(Read-MoveLog $log.FullName | Where-Object { $_.Status -eq 'ПЕРЕНЕСЁН' -and $_.Target }) }
    catch { [void][LibrarianApp.LibBox]::Show("Не удалось прочитать журнал:`r`n$($log.FullName)`r`n$($_.Exception.Message)"); return }
    if ($entries.Count -eq 0) { Rename-Item -LiteralPath $log.FullName -NewName ($log.BaseName + '.undone.csv'); [void][LibrarianApp.LibBox]::Show('В последнем журнале нет перенесённых файлов. Он помечен как отменённый.'); return }
    $when = $log.LastWriteTime.ToString('dd.MM.yyyy HH:mm')
    if ([LibrarianApp.LibBox]::Show("Вернуть $($entries.Count) файлов переноса от $when на прежние места?`r`nЕсли на прежнем месте уже появился файл с тем же именем, этот файл будет пропущен.", 'Отмена переноса', 'YesNo', 'Question') -ne 'Yes') { return }
    $byPath = @{}
    foreach ($entry in $script:AllFiles) { $byPath[[string]$entry.Path] = $entry }
    $returned = 0; $problems = New-Object System.Collections.Generic.List[string]
    [array]::Reverse($entries)
    foreach ($item in $entries) {
        try {
            if (!(Test-Path -LiteralPath $item.Target -PathType Leaf)) {
                if (Test-Path -LiteralPath $item.Source -PathType Leaf) { continue }  # already returned earlier
                $problems.Add("Нет файла на новом месте: $($item.Target)"); continue
            }
            if (Test-Path -LiteralPath $item.Source) { $problems.Add("На прежнем месте уже есть файл: $($item.Source)"); continue }
            [void][IO.Directory]::CreateDirectory([IO.Path]::GetDirectoryName($item.Source))
            Move-Item -LiteralPath $item.Target -Destination $item.Source -ErrorAction Stop
            $returned++
            $entry = $byPath[[string]$item.Target]
            if ($entry) { $entry.Path = $item.Source; $entry.Name = [IO.Path]::GetFileName($item.Source) }
        } catch { $problems.Add("$($item.Target) : $($_.Exception.Message)") }
    }
    $script:AllFiles | Export-Csv -LiteralPath $script:IndexPath -NoTypeInformation -Encoding UTF8
    Reset-DuplicateScan
    Apply-Search
    $resultPath = Join-Path $script:MoveLogDir ($log.BaseName + '_undo_result.txt')
    @("Возвращено: $returned", "Проблем: $($problems.Count)", '') + @($problems) | Set-Content -LiteralPath $resultPath -Encoding UTF8
    # Mark the log as done only when everything came back; otherwise the user can fix the cause and retry.
    if ($problems.Count -eq 0) { Rename-Item -LiteralPath $log.FullName -NewName ($log.BaseName + '.undone.csv') }
    $status.Text = "Возвращено файлов: $returned; проблем: $($problems.Count)."
    $detail = if ($problems.Count -gt 0) { "`r`nПодробности: $resultPath`r`nПосле исправления можно нажать «Отменить перенос» ещё раз." } else { "`r`nПустые папки на книжных полках остались — их можно удалить вручную." }
    [void][LibrarianApp.LibBox]::Show("Возвращено: $returned`r`nПроблем: $($problems.Count)$detail", 'Отмена переноса')
}
function Open-SelectedFile {
    $paths = @(Get-SelectedPaths)
    if ($paths.Count -eq 0) { [void][LibrarianApp.LibBox]::Show('Сначала щёлкни по строке файла. Его полный путь появится в строке над кнопками.'); return }
    if (!(Test-Path -LiteralPath $paths[0] -PathType Leaf)) { [void][LibrarianApp.LibBox]::Show("Файл не найден по этому пути:`r`n$($paths[0])`r`nОбнови каталог и проверь результат."); return }
    try { Start-Process -FilePath $paths[0] -ErrorAction Stop } catch { [void][LibrarianApp.LibBox]::Show("Не удалось открыть файл:`r`n$($paths[0])`r`n`r`nПричина: $($_.Exception.Message)") }
}
function Show-SelectedFileInFolder {
    $paths = @(Get-SelectedPaths)
    if ($paths.Count -eq 0) { [void][LibrarianApp.LibBox]::Show('Сначала щёлкни по строке файла. Его полный путь появится в строке над кнопками.'); return }
    if (!(Test-Path -LiteralPath $paths[0] -PathType Leaf)) { [void][LibrarianApp.LibBox]::Show("Файл не найден по этому пути:`r`n$($paths[0])`r`nОбнови каталог и проверь результат."); return }
    try { Start-Process explorer.exe -ArgumentList "/select,`"$($paths[0])`"" -ErrorAction Stop } catch { [void][LibrarianApp.LibBox]::Show("Не удалось показать файл в папке:`r`n$($paths[0])`r`n`r`nПричина: $($_.Exception.Message)") }
}
function Reset-DuplicateScan {
    $script:HashScanComplete = $false
    $script:DuplicateGroups = @()
    $cleanupButton.Enabled = $false
    foreach ($entry in $script:AllFiles) {
        if ($entry.PSObject.Properties['Hash']) { $entry.Hash = '' }
    }
}
function Get-KeeperFirst($Entries) {
    $downloads = Join-Path $env:USERPROFILE 'Downloads'
    @($Entries | Sort-Object `
        @{ Expression = { if ($_.Path.StartsWith(($downloads.TrimEnd('\') + '\'), [StringComparison]::OrdinalIgnoreCase)) { 1 } else { 0 } } }, `
        @{ Expression = { $_.Path.Length } }, `
        @{ Expression = { $_.Path } })
}
function Get-DocumentPreview([string]$Path, [string]$Extension) {
    $content = ''
    $script:LastContentReadError = ''
    try {
        if ((Get-Item -LiteralPath $Path -ErrorAction Stop).Length -gt 52428800) {
            $script:LastContentReadError = 'Пропущен: файл больше 50 МБ.'
            return ''
        }
        if ($Extension -in @('.txt','.md')) {
            $reader = [System.IO.StreamReader]::new($Path, [System.Text.Encoding]::Default, $true)
            try { $content = $reader.ReadToEnd() } finally { $reader.Dispose() }
        } elseif ($Extension -eq '.docx') {
            try {
                $zip = [System.IO.Compression.ZipFile]::OpenRead($Path)
                try {
                    $entry = $zip.GetEntry('word/document.xml')
                    if ($entry -and $entry.Length -le 52428800) {
                        $stream = $entry.Open()
                        try {
                            $settings = New-Object System.Xml.XmlReaderSettings
                            $settings.DtdProcessing = [System.Xml.DtdProcessing]::Prohibit
                            $xmlReader = [System.Xml.XmlReader]::Create($stream, $settings)
                            try {
                                $xmlDoc = New-Object System.Xml.XmlDocument
                                $xmlDoc.XmlResolver = $null
                                $xmlDoc.Load($xmlReader)
                                $ns = New-Object System.Xml.XmlNamespaceManager($xmlDoc.NameTable)
                                $ns.AddNamespace('w','http://schemas.openxmlformats.org/wordprocessingml/2006/main')
                                $nodes = $xmlDoc.SelectNodes('//w:t', $ns)
                                $content = (@($nodes | ForEach-Object { $_.InnerText }) -join ' ')
                            } finally { $xmlReader.Dispose() }
                        } finally { $stream.Dispose() }
                    } else {
                        $script:LastContentReadError = 'Внутри файла не найден обычный текст Word.'
                    }
                } finally { $zip.Dispose() }
            } catch {
                $script:LastContentReadError = 'Прямое чтение DOCX: ' + $_.Exception.Message
            }
            $content = [System.Text.RegularExpressions.Regex]::Replace([string]$content, '\s+', ' ').Trim()
            if (!$content -and $script:WordApp) {
                $doc = $null
                try {
                    $doc = $script:WordApp.Documents.Open($Path, $false, $true, $false)
                    $content = [string]$doc.Content.Text
                    if ($content) { $script:LastContentReadError = '' }
                } catch {
                    $script:LastContentReadError = 'Word не смог прочитать DOCX: ' + $_.Exception.Message
                } finally {
                    if ($doc) { try { $doc.Close(0) } catch { }; try { [void][Runtime.InteropServices.Marshal]::ReleaseComObject($doc) } catch { } }
                }
            }
        } elseif ($Extension -in @('.doc','.rtf','.pdf')) {
            if ($Extension -eq '.doc' -and !$script:WordMacrosDisabled) { throw 'Не удалось подтвердить отключение макросов Microsoft Word.' }
            if (!$script:WordApp) { throw 'Microsoft Word не запущен для чтения этого формата.' }
            $doc = $null
            try {
                $doc = $script:WordApp.Documents.Open($Path, $false, $true, $false)
                $content = [string]$doc.Content.Text
            } finally {
                if ($doc) { try { $doc.Close(0) } catch { }; try { [void][Runtime.InteropServices.Marshal]::ReleaseComObject($doc) } catch { } }
            }
        }
        $content = [System.Text.RegularExpressions.Regex]::Replace([string]$content, '\s+', ' ').Trim()
        if ($content.Length -gt 1200) { $content = $content.Substring(0,1200) }
        if (!$content -and !$script:LastContentReadError) {
            $script:LastContentReadError = 'Текст не найден. Возможно, документ состоит из изображений или сканов; распознавания текста (OCR) пока нет.'
        }
        return $content
    } catch {
        $script:ContentReadErrors++
        $script:LastContentReadError = $_.Exception.Message
        return ''
    }
}

# ---------------- Еженедельная уборка «Загрузок» и Рабочего стола ----------------
function Get-DownloadsPath {
    try {
        $userShell = Get-ItemProperty -LiteralPath 'HKCU:\Software\Microsoft\Windows\CurrentVersion\Explorer\User Shell Folders' -ErrorAction Stop
        $path = [Environment]::ExpandEnvironmentVariables([string]$userShell.'{374DE290-123F-4565-9164-39C4925E467B}')
        if ($path -and (Test-Path -LiteralPath $path -PathType Container)) { return $path }
    } catch { }
    return (Join-Path $env:USERPROFILE 'Downloads')
}
function Test-SensitiveName([string]$Text) {
    return ($Text.ToLowerInvariant() -match 'парол|password|passport|паспорт|(^|[^а-яё-])банк|снилс|(^|[^а-яё])инн([^а-яё]|$)')
}
function Get-EntryTime($Entry) {
    try { return [datetime]::ParseExact([string]$Entry.Modified, 'yyyy-MM-dd HH:mm', $null) } catch { return (Get-Date) }
}
function Show-CleanupIntro {
    $dlg = New-Object System.Windows.Forms.Form
    $dlg.Text = 'Библиотекарь: еженедельная уборка'; $dlg.StartPosition = 'CenterScreen'; $dlg.Font = $form.Font
    $dlg.Size = [System.Drawing.Size]::new(660,600); $dlg.FormBorderStyle = 'FixedDialog'; $dlg.MaximizeBox = $false; $dlg.MinimizeBox = $false
    $dlg.TopMost = $true; $dlg.BackColor = [System.Drawing.Color]::FromArgb(250,248,243)
    $pic = New-Object System.Windows.Forms.PictureBox
    $pic.Location = [System.Drawing.Point]::new(12,12); $pic.Size = [System.Drawing.Size]::new(620,340); $pic.SizeMode = 'Zoom'
    # The sweeping scene itself, moving (a PictureBox plays an animated GIF by itself); the still picture is the fallback.
    $sweepGif = @((Join-Path $script:MediaDir 'helper_full\sweep.gif'), (Join-Path $script:MediaDir 'helper\sweep.gif'), $script:SweepImage) | Where-Object { Test-Path -LiteralPath $_ } | Select-Object -First 1
    if ($sweepGif) { try { $pic.Image = [System.Drawing.Image]::FromStream((New-Object System.IO.MemoryStream (,[IO.File]::ReadAllBytes($sweepGif)))) } catch { } }
    $dlg.Add_Shown({ Play-Voice (Join-Path $script:MediaDir 'helper\sweep.wav') })
    $dlg.Controls.Add($pic)
    $quote = New-Object System.Windows.Forms.Label
    $quote.Text = '«Так, наведём порядок! Ненужное — в корзину, нужное — на полочку.»'
    $quote.Font = [System.Drawing.Font]::new('Segoe UI', 11, [System.Drawing.FontStyle]::Italic)
    $quote.Location = [System.Drawing.Point]::new(12,360); $quote.Size = [System.Drawing.Size]::new(620,26)
    $dlg.Controls.Add($quote)
    $text = New-Object System.Windows.Forms.Label
    $text.Text = "Я посмотрю «Загрузки» и Рабочий стол.`r`n• Сама: установщики программ старше $($script:InstallerAgeDays) дней и точные копии — в Корзину; файлы с понятной темой — на книжные полки.`r`n• Про остальное спрошу: отметите галочками, что разложить.`r`n• Не трогаю: файлы младше $($script:CleanupAgeDays) дней, ярлыки и всё, похожее на пароли и документы.`r`nПереносы можно отменить кнопкой «Отменить перенос», удалённое лежит в Корзине."
    $text.Location = [System.Drawing.Point]::new(12,390); $text.Size = [System.Drawing.Size]::new(620,110)
    $dlg.Controls.Add($text)
    $start = New-Object System.Windows.Forms.Button
    $start.Text = 'Начать уборку'; $start.Location = [System.Drawing.Point]::new(12,510); $start.Size = [System.Drawing.Size]::new(170,34)
    $start.DialogResult = [System.Windows.Forms.DialogResult]::OK
    $video = New-Object System.Windows.Forms.Button
    $video.Text = 'Посмотреть мультик'; $video.Location = [System.Drawing.Point]::new(190,510); $video.Size = [System.Drawing.Size]::new(170,34)
    $video.Add_Click({ if (Test-Path -LiteralPath $script:SweepVideo) { try { Start-Process -FilePath $script:SweepVideo } catch { } } })
    $later = New-Object System.Windows.Forms.Button
    $later.Text = 'Не сейчас'; $later.Location = [System.Drawing.Point]::new(512,510); $later.Size = [System.Drawing.Size]::new(120,34)
    $later.DialogResult = [System.Windows.Forms.DialogResult]::Cancel
    Set-WoodButton $start; Set-WoodButton $video -Light; Set-WoodButton $later -Light; $dlg.BackColor = $script:Paper; $quote.ForeColor = $script:Wood; $text.ForeColor = $script:Ink
    $dlg.Controls.Add($start); $dlg.Controls.Add($video); $dlg.Controls.Add($later)
    $dlg.AcceptButton = $start; $dlg.CancelButton = $later
    try { return ($dlg.ShowDialog($form) -eq [System.Windows.Forms.DialogResult]::OK) }
    finally { if ($pic.Image) { $pic.Image.Dispose() }; $dlg.Dispose() }
}
function Get-WeeklyCleanupPlan {
    $downloads = Get-DownloadsPath
    $desktop = [Environment]::GetFolderPath([Environment+SpecialFolder]::Desktop)
    $sources = @(@(@{ Path=$downloads; Label='Загрузки' }, @{ Path=$desktop; Label='Рабочий стол' }) | Where-Object { $_.Path -and (Test-Path -LiteralPath $_.Path -PathType Container) })
    foreach ($s in $sources) { $s.Prefix = ([string]$s.Path).TrimEnd('\') + '\' }
    $prefixes = @($sources | ForEach-Object { $_.Prefix } | Sort-Object Length -Descending)
    $storage = Read-SmallTextFile $script:StorageRootConfigPath
    $storagePrefix = if ($storage) { $storage.TrimEnd('\') + '\' } else { '' }
    $now = Get-Date
    $freshCut = $now.AddDays(-$script:CleanupAgeDays)
    $installerCut = $now.AddDays(-$script:InstallerAgeDays)
    $skipExt = @('.lnk','.url','.crdownload','.part','.partial','.tmp','.download','.opdownload')

    # 1. Files of Downloads/Desktop from the fresh catalog, grouped as the plan groups them (top folder or loose file).
    $items = New-Object System.Collections.Generic.List[object]
    foreach ($entry in $script:AllFiles) {
        $p = [string]$entry.Path
        $src = $sources | Where-Object { $p.StartsWith($_.Prefix, [StringComparison]::OrdinalIgnoreCase) } | Select-Object -First 1
        if (!$src) { continue }
        if ($storagePrefix -and $p.StartsWith($storagePrefix, [StringComparison]::OrdinalIgnoreCase)) { continue }
        $name = [string]$entry.Name
        if ($name -like '~$*' -or $name -eq 'desktop.ini' -or $skipExt -contains ([string]$entry.Extension).ToLowerInvariant()) { continue }
        $relative = $p.Substring($src.Prefix.Length)
        $isLoose = ($relative -notmatch '\\')
        $topKey = if ($isLoose) { '' } else { ($relative -split '\\')[0] }
        $items.Add([pscustomobject]@{ Entry=$entry; Source=$src.Label; Relative=$relative; Loose=$isLoose; Top=$topKey; Time=(Get-EntryTime $entry); Sensitive=(Test-SensitiveName $relative) })
    }

    # A folder counts as fresh if anything inside it changed recently: it is probably still being worked on.
    $freshFolders = @{}
    foreach ($it in $items) { if (!$it.Loose -and $it.Time -gt $freshCut) { $freshFolders[$it.Source + '|' + $it.Top] = $true } }

    $trash = New-Object System.Collections.Generic.List[object]
    $work = New-Object System.Collections.Generic.List[object]
    $fresh = 0; $sensitive = New-Object System.Collections.Generic.List[string]
    $installerExt = @('.exe','.msi','.msix','.msixbundle','.appx','.appxbundle')
    foreach ($it in $items) {
        if ($it.Sensitive) { $sensitive.Add([string]$it.Entry.Path); continue }
        if ($it.Loose) { if ($it.Time -gt $freshCut) { $fresh++; continue } }
        elseif ($freshFolders.ContainsKey($it.Source + '|' + $it.Top)) { $fresh++; continue }
        if ($it.Loose -and $it.Source -eq 'Загрузки' -and $installerExt -contains ([string]$it.Entry.Extension).ToLowerInvariant() -and $it.Time -lt $installerCut) {
            $trash.Add([pscustomobject]@{ Path=[string]$it.Entry.Path; Reason='установщик программы старше ' + $script:InstallerAgeDays + ' дней'; Hash=''; Keeper=''; SizeBytes=[long]$it.Entry.SizeBytes }); continue
        }
        $work.Add($it)
    }

    # 2. Exact copies (SHA-256) among loose files: the copy in Downloads/Desktop goes to the Recycle Bin, one copy stays.
    $status.Text = 'Ищу точные копии…'; $form.Refresh(); [Windows.Forms.Application]::DoEvents()
    $looseBySize = @{}
    foreach ($it in $work) { if ($it.Loose -and [long]$it.Entry.SizeBytes -gt 0) { $looseBySize[[long]$it.Entry.SizeBytes] = $true } }
    $hashGroups = @{}
    $hashCache = @{}
    foreach ($group in @($script:AllFiles | Where-Object { $looseBySize.ContainsKey([long]$_.SizeBytes) } | Group-Object { [long]$_.SizeBytes })) {
        if ($group.Count -lt 2) { continue }
        foreach ($entry in $group.Group) {
            try {
                if (!(Test-Path -LiteralPath $entry.Path -PathType Leaf)) { continue }
                $h = (Get-FileHash -LiteralPath $entry.Path -Algorithm SHA256 -ErrorAction Stop).Hash
                $hashCache[[string]$entry.Path] = $h
                if (!$hashGroups.ContainsKey($h)) { $hashGroups[$h] = New-Object System.Collections.Generic.List[object] }
                $hashGroups[$h].Add($entry)
            } catch { }
        }
        [Windows.Forms.Application]::DoEvents()
    }
    $dupPaths = @{}
    $looseWork = @{}
    foreach ($it in $work) { if ($it.Loose) { $looseWork[[string]$it.Entry.Path] = $true } }
    foreach ($pair in $hashGroups.GetEnumerator()) {
        if ($pair.Value.Count -lt 2) { continue }
        $ordered = @(Get-KeeperFirst $pair.Value)
        $keeper = [string]$ordered[0].Path
        foreach ($other in @($ordered | Select-Object -Skip 1)) {
            $op = [string]$other.Path
            if ($looseWork.ContainsKey($op)) {
                $dupPaths[$op] = $true
                $trash.Add([pscustomobject]@{ Path=$op; Reason='точная копия файла ' + $keeper; Hash=$pair.Key; Keeper=$keeper; SizeBytes=[long]$other.SizeBytes })
            }
        }
    }

    # 3. Everything else: confident rows go to the shelves by themselves, the rest waits for a tick.
    $rows = New-Object System.Collections.Generic.List[object]
    foreach ($it in $work) {
        if ($dupPaths.ContainsKey([string]$it.Entry.Path)) { continue }
        $row = Get-PlanRow $it.Entry $prefixes
        $key = if ($it.Loose) { $it.Source + ': файлы ' + (Get-TypeFolder ([string]$it.Entry.Extension)) } else { $it.Source + ': папка «' + $it.Top + '»' }
        $row | Add-Member -NotePropertyName 'Ключ' -NotePropertyValue $key -Force
        $rows.Add($row)
    }
    $sure = @($rows | Where-Object { $_.'Перенести' -eq 'да' })
    $unsure = @($rows | Where-Object { $_.'Перенести' -ne 'да' })
    $groups = @($unsure | Group-Object 'Ключ' | Sort-Object Count -Descending | ForEach-Object {
        $first = $_.Group | Sort-Object { ([string]$_.'ПредлагаемаяПапка').Length } | Select-Object -First 1
        $target = [string]$first.'ПредлагаемаяПапка'
        if ($_.Name -like '*: папка*') { $target = [string]$first.'КудаГруппа' }
        [pscustomobject]@{ Name=$_.Name; Count=$_.Count; Target=$target; Rows=@($_.Group) }
    })
    return [pscustomobject]@{ Trash=$trash; Sure=$sure; UnsureGroups=$groups; Fresh=$fresh; Sensitive=$sensitive; Downloads=$downloads }
}
function Show-CleanupPlanDialog($Plan) {
    $trashBytes = [long](($Plan.Trash | Measure-Object SizeBytes -Sum).Sum)
    $installers = @($Plan.Trash | Where-Object { !$_.Hash }).Count
    $copies = @($Plan.Trash | Where-Object { $_.Hash }).Count
    $sureFolders = @($Plan.Sure | Group-Object 'Ключ' | Sort-Object Count -Descending | Select-Object -First 6 | ForEach-Object { '      ' + $_.Name + ' — ' + $_.Count })
    $dlg = New-Object System.Windows.Forms.Form
    $dlg.Text = 'План уборки'; $dlg.StartPosition = 'CenterScreen'; $dlg.Font = $form.Font; $dlg.TopMost = $true
    $dlg.Size = [System.Drawing.Size]::new(860,660); $dlg.FormBorderStyle = 'FixedDialog'; $dlg.MinimizeBox = $false; $dlg.MaximizeBox = $false
    $info = New-Object System.Windows.Forms.Label
    $auto = "СДЕЛАЮ САМА:`r`n  • в Корзину: $($Plan.Trash.Count) (установщики: $installers, точные копии: $copies) — освободится $(Format-Bytes $trashBytes)`r`n  • на книжные полки по темам: $($Plan.Sure.Count) файлов"
    if ($sureFolders.Count -gt 0) { $auto += "`r`n" + ($sureFolders -join "`r`n") }
    $auto += "`r`nНЕ ТРОГАЮ: свежих (младше $($script:CleanupAgeDays) дней): $($Plan.Fresh); похожих на пароли и документы: $($Plan.Sensitive.Count).`r`n`r`nОТМЕТЬТЕ, ЧТО ЕЩЁ РАЗЛОЖИТЬ (неотмеченное останется на месте до следующей уборки):"
    $info.Text = $auto
    $info.Location = [System.Drawing.Point]::new(12,10); $info.Size = [System.Drawing.Size]::new(820,190)
    $dlg.Controls.Add($info)
    $list = New-Object System.Windows.Forms.CheckedListBox
    $list.CheckOnClick = $true; $list.HorizontalScrollbar = $true
    $list.Location = [System.Drawing.Point]::new(12,205); $list.Size = [System.Drawing.Size]::new(820,345)
    foreach ($g in $Plan.UnsureGroups) { [void]$list.Items.Add(('{0} — {1} файлов → {2}' -f $g.Name, $g.Count, $g.Target)) }
    $dlg.Controls.Add($list)
    $all = New-Object System.Windows.Forms.Button
    $all.Text = 'Отметить всё'; $all.Location = [System.Drawing.Point]::new(12,565); $all.Size = [System.Drawing.Size]::new(130,32)
    $all.Add_Click({ for ($k = 0; $k -lt $list.Items.Count; $k++) { $list.SetItemChecked($k, $true) } })
    $details = New-Object System.Windows.Forms.Button
    $details.Text = 'Подробный список'; $details.Location = [System.Drawing.Point]::new(150,565); $details.Size = [System.Drawing.Size]::new(150,32)
    $details.Add_Click({
        $lines = New-Object System.Collections.Generic.List[string]
        $lines.Add('ПЛАН ЕЖЕНЕДЕЛЬНОЙ УБОРКИ (ничего ещё не сделано)'); $lines.Add('')
        $lines.Add('В КОРЗИНУ:'); foreach ($t in $Plan.Trash) { $lines.Add('  ' + $t.Path + '   (' + $t.Reason + ')') }
        $lines.Add(''); $lines.Add('НА ПОЛКИ САМА:'); foreach ($r in $Plan.Sure) { $lines.Add('  ' + $r.'ИсходныйПуть' + '  →  ' + $r.'ПредлагаемаяПапка') }
        $lines.Add(''); $lines.Add('ЖДУТ ВАШЕЙ ГАЛОЧКИ:'); foreach ($g in $Plan.UnsureGroups) { $lines.Add('  [' + $g.Name + ']'); foreach ($r in $g.Rows) { $lines.Add('    ' + $r.'ИсходныйПуть' + '  →  ' + $r.'ПредлагаемаяПапка') } }
        $lines.Add(''); $lines.Add('НЕ ТРОГАЮ (похоже на личные данные):'); foreach ($s in $Plan.Sensitive) { $lines.Add('  ' + $s) }
        [void][IO.Directory]::CreateDirectory($script:WeeklyReportDir)
        $p = Join-Path $script:WeeklyReportDir 'plan_preview.txt'
        $lines | Set-Content -LiteralPath $p -Encoding UTF8
        Start-Process notepad.exe -ArgumentList ('"' + $p + '"')
    })
    $ok = New-Object System.Windows.Forms.Button
    $ok.Text = 'Убрать'; $ok.Location = [System.Drawing.Point]::new(600,565); $ok.Size = [System.Drawing.Size]::new(120,32)
    $ok.DialogResult = [System.Windows.Forms.DialogResult]::OK
    $cancel = New-Object System.Windows.Forms.Button
    $cancel.Text = 'Отмена'; $cancel.Location = [System.Drawing.Point]::new(728,565); $cancel.Size = [System.Drawing.Size]::new(104,32)
    $cancel.DialogResult = [System.Windows.Forms.DialogResult]::Cancel
    $dlg.Controls.AddRange(@($all,$details,$ok,$cancel))
    $dlg.AcceptButton = $ok; $dlg.CancelButton = $cancel
    try {
        if ($dlg.ShowDialog($form) -ne [System.Windows.Forms.DialogResult]::OK) { return @{ Ok=$false; Rows=@() } }
        $picked = New-Object System.Collections.Generic.List[object]
        foreach ($i in $list.CheckedIndices) { foreach ($r in $Plan.UnsureGroups[$i].Rows) { $picked.Add($r) } }
        return @{ Ok=$true; Rows=$picked.ToArray() }
    } finally { $dlg.Dispose() }
}
function Send-ToRecycleBin($Item, [string]$LogPath) {
    # Copies are re-checked right before deletion: the kept file must still exist with the same content.
    if (!(Test-Path -LiteralPath $Item.Path -PathType Leaf)) { Add-MoveLogEntry $LogPath 'ПРОПУЩЕН' $Item.Path '' 'Файл не найден.'; return $false }
    if ($Item.Hash) {
        if (!(Test-Path -LiteralPath $Item.Keeper -PathType Leaf)) { Add-MoveLogEntry $LogPath 'ПРОПУЩЕН' $Item.Path '' 'Сохраняемая копия недоступна.'; return $false }
        if ((Get-FileHash -LiteralPath $Item.Path -Algorithm SHA256).Hash -ne $Item.Hash -or (Get-FileHash -LiteralPath $Item.Keeper -Algorithm SHA256).Hash -ne $Item.Hash) {
            Add-MoveLogEntry $LogPath 'ПРОПУЩЕН' $Item.Path '' 'Содержимое изменилось.'; return $false
        }
    }
    [Microsoft.VisualBasic.FileIO.FileSystem]::DeleteFile($Item.Path, [Microsoft.VisualBasic.FileIO.UIOption]::OnlyErrorDialogs, [Microsoft.VisualBasic.FileIO.RecycleOption]::SendToRecycleBin, [Microsoft.VisualBasic.FileIO.UICancelOption]::ThrowException)
    Add-MoveLogEntry $LogPath 'В КОРЗИНЕ' $Item.Path '' $Item.Reason
    return $true
}
function Remove-EmptiedFolders($SourcePaths, $Roots) {
    # Only folders that this cleanup emptied go to the Recycle Bin (walking up, never the root itself).
    $rootSet = @($Roots | ForEach-Object { ([string]$_).TrimEnd('\') })
    $dirs = @{}
    foreach ($p in $SourcePaths) {
        $d = [IO.Path]::GetDirectoryName([string]$p)
        while ($d -and !($rootSet -contains $d.TrimEnd('\'))) {
            if (!(@($rootSet | Where-Object { $d.StartsWith($_ + '\', [StringComparison]::OrdinalIgnoreCase) }).Count)) { break }
            $dirs[$d] = $true
            $d = [IO.Path]::GetDirectoryName($d)
        }
    }
    $removed = 0
    foreach ($d in @($dirs.Keys | Sort-Object Length -Descending)) {
        try {
            if ((Test-Path -LiteralPath $d -PathType Container) -and @(Get-ChildItem -LiteralPath $d -Force -ErrorAction Stop).Count -eq 0) {
                [Microsoft.VisualBasic.FileIO.FileSystem]::DeleteDirectory($d, [Microsoft.VisualBasic.FileIO.UIOption]::OnlyErrorDialogs, [Microsoft.VisualBasic.FileIO.RecycleOption]::SendToRecycleBin, [Microsoft.VisualBasic.FileIO.UICancelOption]::ThrowException)
                $removed++
            }
        } catch { }
    }
    return $removed
}
function Invoke-WeeklyCleanup {
    Show-Helper 'sweep' -Silent
    if (!(Show-CleanupIntro)) { $status.Text = 'Уборка отложена.'; Show-HelperIdle; return }
    Start-HelperLoop 'sweep' -Silent
    $weeklyButton.Enabled = $false
    try {
        Update-Catalog
        $status.Text = 'Составляю план уборки…'; $form.Refresh()
        $plan = Get-WeeklyCleanupPlan
        if ($plan.Trash.Count -eq 0 -and $plan.Sure.Count -eq 0 -and $plan.UnsureGroups.Count -eq 0) {
            $status.Text = 'Убирать нечего.'
            [void][LibrarianApp.LibBox]::Show("В «Загрузках» и на Рабочем столе уже порядок.`r`nСвежих файлов (их пока не трогаю): $($plan.Fresh).", 'Уборка'); return
        }
        $choice = Show-CleanupPlanDialog $plan
        if (!$choice.Ok) { $status.Text = 'Уборка отменена. Ничего не изменено.'; return }

        $storageRoot = Read-SmallTextFile $script:StorageRootConfigPath
        if (!$storageRoot -or !(Test-Path -LiteralPath $storageRoot -PathType Container)) { $storageRoot = Get-StorageRoot; if (!$storageRoot) { return } }
        [void][IO.Directory]::CreateDirectory($script:MoveLogDir)
        $stamp = Get-Date -Format 'yyyyMMdd_HHmmss'
        $logPath = Join-Path $script:MoveLogDir ('moves_' + $stamp + '_уборка.csv')

        $binned = 0; $binBytes = [long]0; $binFailed = 0; $i = 0
        foreach ($t in $plan.Trash) {
            $i++
            try { if (Send-ToRecycleBin $t $logPath) { $binned++; $binBytes += [long]$t.SizeBytes } }
            catch { $binFailed++; Add-MoveLogEntry $logPath 'ОШИБКА' $t.Path '' $_.Exception.Message }
            if (($i % 10) -eq 0) { $status.Text = "В Корзину: $i из $($plan.Trash.Count)…"; $form.Refresh(); [Windows.Forms.Application]::DoEvents() }
        }
        $rows = @(@($plan.Sure) + @($choice.Rows))
        $prepared = Get-MoveTasks $rows $storageRoot
        foreach ($problem in $prepared.Problems) { Add-MoveLogEntry $logPath 'ПРОПУЩЕН' $problem.Source '' $problem.Note }
        $result = Invoke-MoveTasks $prepared.Tasks $logPath
        $touched = @(@($prepared.Tasks | ForEach-Object { $_.Source }) + @($plan.Trash | ForEach-Object { $_.Path }))
        $emptied = Remove-EmptiedFolders $touched @($plan.Downloads, [Environment]::GetFolderPath([Environment+SpecialFolder]::Desktop))

        $script:AllFiles = @($script:AllFiles | Where-Object { Test-Path -LiteralPath $_.Path -PathType Leaf })
        $script:AllFiles | Export-Csv -LiteralPath $script:IndexPath -NoTypeInformation -Encoding UTF8
        Reset-DuplicateScan; Apply-Search
        [void][IO.Directory]::CreateDirectory($script:WeeklyReportDir)
        Set-Content -LiteralPath (Join-Path $script:WeeklyReportDir 'last_run.txt') -Value (Get-Date -Format 'yyyy-MM-dd HH:mm') -Encoding UTF8

        $left = @($choice.Rows).Count
        $notPicked = (@($plan.UnsureGroups | ForEach-Object { $_.Count }) | Measure-Object -Sum).Sum - $left
        $avast = ''
        if (@($result.Errors | Where-Object { $_ -match 'доступ|denied|access' }).Count -gt 0 -or $binFailed -gt 0) {
            $avast = "`r`n`r`nЧасть файлов не сдвинулась: Windows или антивирус не дал доступ. Если включён щит Avast «Защита от программ-вымогателей», выключите его на время уборки и запустите её ещё раз."
        }
        $status.Text = "Уборка: в Корзину $binned, на полки $($result.Moved), ошибок $($result.Failed + $binFailed)."
        Show-Helper 'dust'
        [void][LibrarianApp.LibBox]::Show(("Готово!`r`n`r`nВ Корзину: $binned (освобождено $(Format-Bytes $binBytes))`r`nРазложено по книжным полкам: $($result.Moved)`r`nПустых папок убрано: $emptied`r`nОставлено до следующего раза: $notPicked неотмеченных, $($plan.Fresh) свежих, $($plan.Sensitive.Count) личных`r`nОшибок: $($result.Failed + $binFailed)$avast`r`n`r`nВернуть разложенное: кнопка «Отменить перенос». Удалённое лежит в Корзине 14 дней.`r`nЖурнал: $logPath"), 'Уборка завершена')
    } finally { $weeklyButton.Enabled = $true }
}
$weeklyButton.Add_Click({ Invoke-WeeklyCleanup })
$promptsButton.Add_Click({ Invoke-PromptCollect })
$siteButton.Add_Click({ Open-PromptSite })

foreach ($root in (Get-Roots)) { [void]$rootsBox.Items.Add([string]$root) }
if ($rootsBox.Items.Count -eq 0) { foreach ($root in (Get-DefaultRoots)) { [void]$rootsBox.Items.Add([string]$root) } }
Add-StorageRootToCatalog
Load-Index

$addRoot.Add_Click({
    $dialog = New-Object System.Windows.Forms.FolderBrowserDialog
    $dialog.Description = 'Выбери папку, которую нужно включить в каталог. Скрипт только прочитает список файлов.'
    if ($dialog.ShowDialog() -eq 'OK' -and !$rootsBox.Items.Contains($dialog.SelectedPath)) {
        [void]$rootsBox.Items.Add($dialog.SelectedPath)
        Save-Roots @($rootsBox.Items)
        Reset-DuplicateScan
    }
})
$removeRoot.Add_Click({
    if ($rootsBox.SelectedItem) { $rootsBox.Items.Remove($rootsBox.SelectedItem); Save-Roots @($rootsBox.Items); Reset-DuplicateScan }
})
$scanButton.Add_Click({
    if ($rootsBox.Items.Count -eq 0) { [void][LibrarianApp.LibBox]::Show('Сначала добавь хотя бы одну папку.'); return }
    Update-Catalog
})
function Update-Catalog {
    Reset-DuplicateScan
    $scanButton.Enabled = $false; $status.Text = 'Сканирую выбранные папки…'; $form.Refresh()
    $skip = @('AppData','Windows','Program Files','Program Files (x86)','$Recycle.Bin','System Volume Information','node_modules','.git')
    $found = New-Object System.Collections.Generic.List[object]
    foreach ($root in @($rootsBox.Items)) {
        if (!(Test-Path -LiteralPath $root -PathType Container)) { continue }
        Get-ChildItem -LiteralPath $root -File -Recurse -Force -ErrorAction SilentlyContinue | ForEach-Object {
            $file = $_
            if (($file.Attributes -band [IO.FileAttributes]::ReparsePoint) -ne 0) { return }
            $blocked = $false
            foreach ($part in $file.DirectoryName.Split([IO.Path]::DirectorySeparatorChar)) { if ($skip -contains $part) { $blocked = $true; break } }
            if (!$blocked) {
                $ext = $file.Extension.ToLowerInvariant()
                $found.Add([pscustomobject]@{ Name=$file.Name; Category=(Get-Category $ext); Extension=$ext; SizeMB=[math]::Round($file.Length / 1MB, 2); SizeBytes=[long]$file.Length; Created=$file.CreationTime.ToString('yyyy-MM-dd HH:mm'); Modified=$file.LastWriteTime.ToString('yyyy-MM-dd HH:mm'); ContentPreview=''; Path=$file.FullName; Hash='' })
            }
        }
    }
    $script:AllFiles = @($found | Sort-Object Path -Unique)
    $script:AllFiles | Export-Csv -LiteralPath $script:IndexPath -NoTypeInformation -Encoding UTF8
    Apply-Search
    $status.Text = "В каталоге файлов: $($script:AllFiles.Count)"
    $scanButton.Enabled = $true
}
$contentButton.Add_Click({
    $supported = @('.txt','.md','.docx','.doc','.rtf','.pdf')
    $selectedPaths = @(Get-SelectedPaths)
    if ($selectedPaths.Count -gt 0) {
        $targets = @($script:AllFiles | Where-Object { ($_.Extension -in $supported) -and ($selectedPaths -contains $_.Path) })
    } else {
        $targets = @($script:AllFiles | Where-Object { $_.Extension -in $supported })
    }
    if ($targets.Count -eq 0) { [void][LibrarianApp.LibBox]::Show('Выбери в таблице Word, TXT или PDF-файл. Если нужен полный проход, сначала сними выделение строк.'); return }
    $scopeText = if ($selectedPaths.Count -gt 0) { 'Будут прочитаны только выбранные строки.' } else { 'Будут прочитаны все документы поддерживаемых типов.' }
    $answer = [LibrarianApp.LibBox]::Show(("$scopeText Файлов: $($targets.Count). В карточку попадёт начало текста, не более 1 200 символов. Фрагменты сохранятся только в локальном индексе на этом компьютере. Исходные файлы не изменяются. Для DOC/PDF/RTF используется Microsoft Word в режиме чтения; файлы с паролем или ошибками будут пропущены. Продолжить?"), 'Чтение документов', 'YesNo', 'Information')
    if ($answer -ne [Windows.Forms.DialogResult]::Yes) { return }
    $script:ContentReadErrors = 0
    $needsWord = @($targets | Where-Object { $_.Extension -in @('.doc','.docx','.rtf','.pdf') }).Count -gt 0
    if ($needsWord) {
        try {
            $script:WordApp = New-Object -ComObject Word.Application -ErrorAction Stop
            $script:WordApp.Visible = $false
            $script:WordApp.DisplayAlerts = 0
            $script:WordApp.AutomationSecurity = 3
            $script:WordMacrosDisabled = $true
            try { $script:WordApp.Options.UpdateLinksAtOpen = $false } catch { }
        } catch {
            if ($script:WordApp) { try { $script:WordApp.Quit() } catch { }; try { [void][Runtime.InteropServices.Marshal]::ReleaseComObject($script:WordApp) } catch { } }
            $script:WordApp = $null; $script:WordMacrosDisabled = $false
        }
    }
    $contentButton.Enabled = $false; $scanButton.Enabled = $false
    $readCount = 0; $emptyCount = 0; $i = 0
    $readReport = New-Object System.Collections.Generic.List[string]
    try {
        foreach ($entry in $targets) {
            $i++
            if (!$entry.PSObject.Properties['ContentPreview']) { $entry | Add-Member -NotePropertyName ContentPreview -NotePropertyValue '' }
            if (!$entry.PSObject.Properties['ContentStatus']) { $entry | Add-Member -NotePropertyName ContentStatus -NotePropertyValue '' }
            $entry.ContentPreview = Get-DocumentPreview ([string]$entry.Path) ([string]$entry.Extension)
            if ($entry.ContentPreview) {
                $entry.ContentStatus = 'Текст извлечён'
                $readCount++
            } else {
                $entry.ContentStatus = $script:LastContentReadError
                $emptyCount++
                $readReport.Add(("{0}`r`n  {1}`r`n  Причина: {2}" -f $entry.Name,$entry.Path,$entry.ContentStatus))
            }
            if (($i % 5) -eq 0 -or $i -eq $targets.Count) {
                $status.Text = "Читаю документы: $i из $($targets.Count)…"; $form.Refresh(); [Windows.Forms.Application]::DoEvents()
            }
        }
        $script:AllFiles | Export-Csv -LiteralPath $script:IndexPath -NoTypeInformation -Encoding UTF8
        if ($readReport.Count -gt 0) {
            $readReport | Set-Content -LiteralPath $script:ContentReadReportPath -Encoding UTF8
        } elseif (Test-Path -LiteralPath $script:ContentReadReportPath) {
            Remove-Item -LiteralPath $script:ContentReadReportPath -Force
        }
        Apply-Search
        $status.Text = "Текст доступен для поиска у $readCount документов; без извлечённого текста: $emptyCount. Фрагменты сохранены локально."
        $detail = if ($emptyCount -gt 0) { "`r`nПричины по файлам записаны в таблицу в колонку «Результат чтения» и в отчёт: $script:ContentReadReportPath" } else { '' }
        [void][LibrarianApp.LibBox]::Show(("Готово. Текст извлечён у $readCount файлов; без текста: $emptyCount. Для DOCX сначала используется безопасное чтение текста, затем при необходимости Word в режиме чтения. В локальном каталоге хранится начало текста до 1 200 символов. Исходные документы не меняются.$detail"), 'Индексирование документов', 'OK', 'Information')
    } finally {
        if ($script:WordApp) { try { $script:WordApp.Quit() } catch { }; try { [void][Runtime.InteropServices.Marshal]::ReleaseComObject($script:WordApp) } catch { }; $script:WordApp = $null }; $script:WordMacrosDisabled = $false
        $contentButton.Enabled = $true; $scanButton.Enabled = $true
    }
})
# ---------------- Задания: «вот папка …, её можно удалить», «открой …», «сделай уборку» ----------------
function Get-PathFromText([string]$Text) {
    # A path copied from Explorer: quoted, a file:/// link, or bare «C:\…» (possibly followed by more words).
    $t = $Text -replace 'file:///', ''
    $m = [regex]::Match($t, '"([A-Za-z]:[\\/][^"]*)"')
    if ($m.Success) { $p = $m.Groups[1].Value.Replace('/','\').TrimEnd('\'); if (Test-Path -LiteralPath $p) { return $p } }
    $m = [regex]::Match($t, '[A-Za-z]:[\\/]')
    if (!$m.Success) { return $null }
    $rest = $t.Substring($m.Index).Replace('/','\')
    try { $rest = [Uri]::UnescapeDataString($rest) } catch { }
    $words = @($rest -split ' ')
    # Longest prefix that exists on disk: the words after the path are the instruction.
    for ($n = $words.Count; $n -ge 1; $n--) {
        $p = (($words[0..($n - 1)]) -join ' ').Trim().TrimEnd(',', '.', ';', ':', '!', '?', '»', '"', "'").TrimEnd('\')
        if ($p -match '^[A-Za-z]:$') { $p += '\' }
        if ($p -and (Test-Path -LiteralPath $p)) { return $p }
    }
    return ('?' + $words[0])   # looks like a path, but nothing is there
}
function Test-ProtectedPath([string]$Path) {
    # Never offered for deletion: system folders, the user's main folders themselves, the shelves and the program.
    $full = [IO.Path]::GetFullPath($Path).TrimEnd('\')
    if ($full -match '^[A-Za-z]:$') { return $true }
    $guard = New-Object System.Collections.Generic.List[string]
    foreach ($p in @($env:SystemRoot, $env:ProgramFiles, ${env:ProgramFiles(x86)}, $env:ProgramData, $env:USERPROFILE, $env:APPDATA, $env:LOCALAPPDATA, (Get-DownloadsPath), $PSScriptRoot, $script:DataHome)) { if ($p) { $guard.Add($p) } }
    foreach ($sf in 'Desktop','MyDocuments','MyPictures','MyVideos','MyMusic') { $p = [Environment]::GetFolderPath($sf); if ($p) { $guard.Add($p) } }
    $storage = Read-SmallTextFile $script:StorageRootConfigPath
    if ($storage) {
        $guard.Add($storage)
        foreach ($shelf in @(Get-ChildItem -LiteralPath $storage -Directory -ErrorAction SilentlyContinue)) { $guard.Add($shelf.FullName) }
    }
    foreach ($g in $guard) {
        $gf = $g.TrimEnd('\')
        # The folder itself, anything above it, or anything inside system/program folders.
        if ([string]::Equals($full, $gf, [StringComparison]::OrdinalIgnoreCase)) { return $true }
        if ($gf.StartsWith($full + '\', [StringComparison]::OrdinalIgnoreCase)) { return $true }
    }
    foreach ($sys in @($env:SystemRoot, $env:ProgramFiles, ${env:ProgramFiles(x86)}, $env:ProgramData, $env:APPDATA, $env:LOCALAPPDATA, $PSScriptRoot)) {
        if ($sys -and $full.StartsWith($sys.TrimEnd('\') + '\', [StringComparison]::OrdinalIgnoreCase)) { return $true }
    }
    return $false
}
function Show-PathChoice([string]$Path, [string]$Summary) {
    $dlg = New-Object System.Windows.Forms.Form
    $dlg.Text = 'Что сделать?'; $dlg.StartPosition = 'CenterParent'; $dlg.Font = $form.Font
    $dlg.Size = [System.Drawing.Size]::new(620,230); $dlg.FormBorderStyle = 'FixedDialog'; $dlg.MinimizeBox = $false; $dlg.MaximizeBox = $false
    $l = New-Object System.Windows.Forms.Label
    $l.Text = $Summary + "`r`n`r`nЧто с этим сделать?"; $l.Location = [System.Drawing.Point]::new(12,12); $l.Size = [System.Drawing.Size]::new(580,110)
    $dlg.Controls.Add($l)
    $choice = @{ Value='cancel' }
    $b1 = New-Object System.Windows.Forms.Button; $b1.Text = 'Открыть'; $b1.Location = [System.Drawing.Point]::new(12,135); $b1.Size = [System.Drawing.Size]::new(140,32)
    $b1.Add_Click({ $choice.Value = 'open'; $dlg.Close() })
    $b2 = New-Object System.Windows.Forms.Button; $b2.Text = 'В Корзину'; $b2.Location = [System.Drawing.Point]::new(160,135); $b2.Size = [System.Drawing.Size]::new(140,32)
    $b2.Add_Click({ $choice.Value = 'delete'; $dlg.Close() })
    $b4 = New-Object System.Windows.Forms.Button; $b4.Text = 'Разобрать картинки'; $b4.Location = [System.Drawing.Point]::new(308,135); $b4.Size = [System.Drawing.Size]::new(150,32)
    $b4.Add_Click({ $choice.Value = 'images'; $dlg.Close() })
    if (!(Test-Path -LiteralPath $Path -PathType Container)) { $b4.Visible = $false }
    $dlg.Controls.Add($b4)
    $b3 = New-Object System.Windows.Forms.Button; $b3.Text = 'Отмена'; $b3.Location = [System.Drawing.Point]::new(472,135); $b3.Size = [System.Drawing.Size]::new(120,32)
    $b3.Add_Click({ $dlg.Close() })
    $dlg.Controls.AddRange(@($b1,$b2,$b3)); $dlg.CancelButton = $b3
    try { [void]$dlg.ShowDialog($form) } finally { $dlg.Dispose() }
    return $choice.Value
}
function Get-PathSummary([string]$Path) {
    if (Test-Path -LiteralPath $Path -PathType Leaf) {
        $f = Get-Item -LiteralPath $Path
        return [pscustomobject]@{ IsFolder=$false; Files=1; Bytes=[long]$f.Length; Sensitive=@($(if (Test-SensitiveName $f.Name) { $f.FullName })); Text=('Файл «' + $f.Name + '», ' + (Format-Bytes $f.Length) + '.') }
    }
    $files = @(Get-ChildItem -LiteralPath $Path -File -Recurse -Force -ErrorAction SilentlyContinue)
    $bytes = [long](($files | Measure-Object Length -Sum).Sum)
    $sens = @($files | Where-Object { Test-SensitiveName $_.FullName.Substring($Path.Length) } | ForEach-Object { $_.FullName })
    return [pscustomobject]@{ IsFolder=$true; Files=$files.Count; Bytes=$bytes; Sensitive=$sens; Text=('Папка «' + [IO.Path]::GetFileName($Path) + '»: файлов ' + $files.Count + ', ' + (Format-Bytes $bytes) + '.') }
}
function Remove-ToRecycleBin([string]$Path) {
    if (Test-ProtectedPath $Path) {
        $answerLabel.Text = 'Эту папку я не трогаю: она системная или главная (Загрузки, Рабочий стол, книжные полки, сама программа). Можно удалить только то, что лежит внутри.'
        return
    }
    $info = Get-PathSummary $Path
    $warn = ''
    if ($info.Sensitive.Count -gt 0) { $warn = "`r`n`r`nВНИМАНИЕ: внутри есть файлы, похожие на пароли или документы ($($info.Sensitive.Count)), например:`r`n" + (($info.Sensitive | Select-Object -First 3 | ForEach-Object { '  ' + [IO.Path]::GetFileName($_) }) -join "`r`n") }
    $big = if ($info.Bytes -gt 2GB) { "`r`n`r`nПапка большая. Если Корзина её не вместит, Windows спросит, удалить ли навсегда, — тогда нажмите «Нет»." } else { '' }
    if ($warn) { Show-Helper 'secret' } else { Show-Helper 'sweep' -Silent }
    $msg = "Отправить в Корзину?`r`n`r`n$($info.Text)`r`n$Path$warn$big`r`n`r`nИз Корзины можно вернуть (она сама очищается через 14 дней)."
    if ([LibrarianApp.LibBox]::Show($msg, 'В Корзину', 'YesNo', $(if ($warn) { 'Warning' } else { 'Question' })) -ne 'Yes') { $answerLabel.Text = 'Хорошо, ничего не удаляю.'; return }
    try {
        # AllDialogs: if something cannot go to the Recycle Bin, Windows asks instead of deleting for good.
        if ($info.IsFolder) {
            [Microsoft.VisualBasic.FileIO.FileSystem]::DeleteDirectory($Path, [Microsoft.VisualBasic.FileIO.UIOption]::AllDialogs, [Microsoft.VisualBasic.FileIO.RecycleOption]::SendToRecycleBin, [Microsoft.VisualBasic.FileIO.UICancelOption]::ThrowException)
        } else {
            [Microsoft.VisualBasic.FileIO.FileSystem]::DeleteFile($Path, [Microsoft.VisualBasic.FileIO.UIOption]::AllDialogs, [Microsoft.VisualBasic.FileIO.RecycleOption]::SendToRecycleBin, [Microsoft.VisualBasic.FileIO.UICancelOption]::ThrowException)
        }
    } catch {
        $answerLabel.Text = 'Не получилось: ' + $_.Exception.Message + ' Если мешает Avast (защита от программ-вымогателей), выключите его щит на минуту и повторите.'
        return
    }
    $prefix = $Path.TrimEnd('\') + '\'
    $script:AllFiles = @($script:AllFiles | Where-Object { !([string]$_.Path).StartsWith($prefix, [StringComparison]::OrdinalIgnoreCase) -and [string]$_.Path -ne $Path })
    $script:AllFiles | Export-Csv -LiteralPath $script:IndexPath -NoTypeInformation -Encoding UTF8
    Reset-DuplicateScan
    Show-Rows $script:AllFiles
    $answerLabel.Text = 'Готово: «' + [IO.Path]::GetFileName($Path) + '» в Корзине (' + $info.Files + ' ' + (Get-FilesWord $info.Files) + ', ' + (Format-Bytes $info.Bytes) + '). Передумали — откройте Корзину и нажмите «Восстановить».'
    $status.Text = 'Удалено в Корзину: ' + $Path
    Show-Helper 'sweep'
}
function Open-PathPlace([string]$Path) {
    if (Test-Path -LiteralPath $Path -PathType Leaf) { Start-Process explorer.exe -ArgumentList "/select,`"$Path`"" } else { Start-Process explorer.exe -ArgumentList "`"$Path`"" }
    $answerLabel.Text = 'Открыл: ' + $Path
}
function Test-LooksLikeCommand([string]$Text) {
    $t = $Text.ToLowerInvariant()
    return ($t -match '[a-z]:[\\/]' -or $t -match '(^|\s)(сделай|запусти|начни)\s+(уборк|порядок)|наведи\s+порядок|обнови\s+каталог|отмени\s+перенос' -or
            $t -match '(удал|выбрось|выброси|выкин|убери|снеси|сотри|откр|покажи)\S*\s+(.*\s)?(папк|файл)' -or
            $t -match '(папк|файл)\S*\s+(.*\s)?(удал|выбрось|выброси|выкин|убер|снес|сотр|откр)')
}
function Get-NameTask([string]$Text) {
    # Both word orders: «удали папку вакансии» and «эту папку вакансии можно удалять».
    $t = $Text.ToLowerInvariant().TrimEnd('.', '!', '?', ' ')
    $m = [regex]::Match($t, '(удал|выбрось|выброси|выкин|убери|снеси|сотри|откр|покажи)\S*\s+(.*?\s)?(папк|файл)\S*\s*(.*)$')
    if ($m.Success) { return @{ Verb=$m.Groups[1].Value; Kind=$m.Groups[3].Value; What=$m.Groups[4].Value } }
    $m = [regex]::Match($t, '^.*(папк|файл)\S*\s*(.*?)\s*(можно\s+|нужно\s+|надо\s+|полностью\s+|целиком\s+)*(удал|выбрось|выброси|выкин|убер|снес|сотр|откр)')
    if ($m.Success) { return @{ Verb=$m.Groups[4].Value; Kind=$m.Groups[1].Value; What=$m.Groups[2].Value } }
    return $null
}
function Find-ByName([string]$Words, [switch]$Folders) {
    # «папку Telegram Desktop» → folders whose own name has all the words (from the beginning of a word).
    $terms = Get-SearchTerms $Words
    $stems = @($terms.Stems)
    if ($stems.Count -eq 0) { return @() }
    $patterns = @($stems | ForEach-Object { New-Object System.Text.RegularExpressions.Regex ('(?<![\p{L}\p{Nd}])' + [regex]::Escape($_)) })
    $rootPrefixes = @(@($rootsBox.Items) | ForEach-Object { ([string]$_).TrimEnd('\') })
    $names = @{}
    foreach ($e in $script:AllFiles) {
        if ($Folders) {
            $d = [IO.Path]::GetDirectoryName([string]$e.Path)
            while ($d -and !$names.ContainsKey($d)) {
                if (@($rootPrefixes | Where-Object { [string]::Equals($_, $d, [StringComparison]::OrdinalIgnoreCase) }).Count) { break }
                $names[$d] = [IO.Path]::GetFileName($d)
                $d = [IO.Path]::GetDirectoryName($d)
            }
        } else { $names[[string]$e.Path] = [string]$e.Name }
    }
    @($names.GetEnumerator() | Where-Object {
        $leaf = ([string]$_.Value).ToLowerInvariant().Replace('ё','е')
        @($patterns | Where-Object { !$_.IsMatch($leaf) }).Count -eq 0
    } | ForEach-Object { [string]$_.Key } | Where-Object { Test-Path -LiteralPath $_ } | Sort-Object Length)
}
function Show-PickList([string]$Title, $Paths) {
    $dlg = New-Object System.Windows.Forms.Form
    $dlg.Text = $Title; $dlg.StartPosition = 'CenterParent'; $dlg.Font = $form.Font
    $dlg.Size = [System.Drawing.Size]::new(760,420); $dlg.FormBorderStyle = 'FixedDialog'; $dlg.MinimizeBox = $false; $dlg.MaximizeBox = $false
    $l = New-Object System.Windows.Forms.Label; $l.Text = 'Нашёл несколько. Выберите нужное:'; $l.Location = [System.Drawing.Point]::new(12,10); $l.AutoSize = $true
    $list = New-Object System.Windows.Forms.ListBox; $list.Location = [System.Drawing.Point]::new(12,34); $list.Size = [System.Drawing.Size]::new(720,290); $list.HorizontalScrollbar = $true
    foreach ($p in $Paths) { [void]$list.Items.Add($p) }
    if ($list.Items.Count) { $list.SelectedIndex = 0 }
    $ok = New-Object System.Windows.Forms.Button; $ok.Text = 'Выбрать'; $ok.Location = [System.Drawing.Point]::new(512,335); $ok.Size = [System.Drawing.Size]::new(110,32); $ok.DialogResult = 'OK'
    $cancel = New-Object System.Windows.Forms.Button; $cancel.Text = 'Отмена'; $cancel.Location = [System.Drawing.Point]::new(628,335); $cancel.Size = [System.Drawing.Size]::new(104,32); $cancel.DialogResult = 'Cancel'
    $list.Add_DoubleClick({ $dlg.DialogResult = 'OK'; $dlg.Close() })
    $dlg.Controls.AddRange(@($l,$list,$ok,$cancel)); $dlg.AcceptButton = $ok; $dlg.CancelButton = $cancel
    try { if ($dlg.ShowDialog($form) -eq 'OK' -and $list.SelectedItem) { return [string]$list.SelectedItem } else { return $null } } finally { $dlg.Dispose() }
}
function Invoke-PathTask([string]$Path, [string]$Intent) {
    switch ($Intent) {
        'delete' { Remove-ToRecycleBin $Path }
        'open' { Open-PathPlace $Path }
        default {
            $info = Get-PathSummary $Path
            $answerLabel.Text = 'Понял: ' + $info.Text + ' Выберите, что сделать.'
            switch (Show-PathChoice $Path ($info.Text + "`r`n" + $Path)) {
                'open' { Open-PathPlace $Path }
                'delete' { Remove-ToRecycleBin $Path }
                'images' { Invoke-ImageCleanup $Path }
                default { $answerLabel.Text = 'Хорошо, ничего не делаю.' }
            }
        }
    }
}
# ---------------- Библиотека промптов: сбор из Telegram-канала и сайт (путь — в настройки.json) ----------------
function Invoke-External([string]$CommandLine, [string]$Dir, [int]$TimeoutSec = 1200) {
    # Runs a console command hidden, keeps the window alive (the Librarian keeps moving), returns its output.
    $out = [IO.Path]::GetTempFileName(); $err = [IO.Path]::GetTempFileName()
    $psi = New-Object System.Diagnostics.ProcessStartInfo
    $psi.FileName = "$env:SystemRoot\System32\cmd.exe"
    $psi.Arguments = '/d /c chcp 65001 >nul & ' + $CommandLine + ' > "' + $out + '" 2> "' + $err + '"'
    $psi.WorkingDirectory = $Dir; $psi.UseShellExecute = $false; $psi.CreateNoWindow = $true
    $psi.EnvironmentVariables['PYTHONIOENCODING'] = 'utf-8'
    $p = [System.Diagnostics.Process]::Start($psi)
    $sw = [Diagnostics.Stopwatch]::StartNew()
    while (!$p.HasExited) {
        [Windows.Forms.Application]::DoEvents(); Start-Sleep -Milliseconds 120
        if ($sw.Elapsed.TotalSeconds -gt $TimeoutSec) { try { $p.Kill() } catch { }; break }
    }
    $p.WaitForExit()
    $o = try { [IO.File]::ReadAllText($out, [Text.Encoding]::UTF8) } catch { '' }
    $e = try { [IO.File]::ReadAllText($err, [Text.Encoding]::UTF8) } catch { '' }
    Remove-Item -LiteralPath $out, $err -ErrorAction SilentlyContinue
    return [pscustomobject]@{ Code=$p.ExitCode; Out=$o.Trim(); Err=$e.Trim() }
}
function Open-PromptSite {
    # The site works only while a small local server runs; start it if it is not running yet.
    $site = Join-Path $script:PromptLibDir 'site'
    if (!(Test-Path -LiteralPath $site)) { $answerLabel.Text = 'Не нашёл папку сайта: ' + $site; return }
    $listening = $false
    try { $c = New-Object Net.Sockets.TcpClient; $c.Connect('127.0.0.1', 8777); $listening = $c.Connected; $c.Close() } catch { }
    if (!$listening) {
        $py = (Get-Command python -ErrorAction SilentlyContinue).Source
        if (!$py) { $answerLabel.Text = 'Не нашёл Python — без него сайт не запустить.'; return }
        Start-Process -FilePath $py -ArgumentList '-m','http.server','8777','--bind','127.0.0.1' -WorkingDirectory $site -WindowStyle Hidden
        Start-Sleep -Milliseconds 1200
    }
    Start-Process 'http://localhost:8777/'
    $answerLabel.Text = 'Открыл сайт библиотеки промптов: http://localhost:8777 (работает, пока включён компьютер).'
}
function Invoke-PromptCollect {
    $lib = $script:PromptLibDir
    if (!(Test-Path -LiteralPath (Join-Path $lib 'scripts\import_bot.py'))) { $answerLabel.Text = 'Не нашёл библиотеку промптов: ' + $lib; return }
    $promptsButton.Enabled = $false
    # He walks along the shelves the whole time the work goes on.
    Start-HelperLoop 'search'
    try {
        $answerLabel.Text = 'Забираю, что бот собрал из канала (GitHub)…'
        $pull = Invoke-External 'git pull --no-edit' $lib 180
        if ($pull.Code -ne 0) { $answerLabel.Text = 'Не получилось забрать обновления с GitHub: ' + ($pull.Err + ' ' + $pull.Out).Trim(); Show-HelperIdle; return }
        $answerLabel.Text = 'Готовлю новые посты к разбору…'
        $imp = Invoke-External 'python scripts\import_bot.py' $lib 300
        $impText = ($imp.Out + ' ' + $imp.Err)
        if ($impText -match 'Новых постов от бота нет|ещё ничего не собрал' -or $imp.Code -ne 0) {
            $answerLabel.Text = 'Новых постов в канале нет — всё уже в библиотеке. ' + $(if ($impText -notmatch 'Новых постов') { $impText.Trim() } else { '' })
            Show-HelperIdle; return
        }
        $answerLabel.Text = 'Разбираю новые посты: что за промпт, раздел, название, теги… Это может занять несколько минут.'
        $claude = (Get-Command claude.cmd -ErrorAction SilentlyContinue).Source
        if (!$claude) { $answerLabel.Text = 'Посты подготовлены (НОВОЕ.md), но Claude Code не найден — разбор сделайте в чате с Claude.'; Show-HelperIdle; return }
        $ask = 'Разбери новые находки строго по инструкции из файла РАЗБОР.md в этой папке. Работай молча, в конце напиши одну строку итога.'
        $cl = Invoke-External ('"' + $claude + '" -p "' + $ask + '" --allowedTools "Read,Write,Edit,Glob,Grep,Bash(python scripts/apply_inbox.py),Bash(python scripts\apply_inbox.py)"') $lib 1500
        $summary = @($cl.Out -split "`r?`n" | Where-Object { $_.Trim() }) | Select-Object -Last 1
        if ($cl.Code -ne 0 -or $summary -notmatch 'Разобрано') {
            $answerLabel.Text = 'Разбор не закончился. Ответ Claude: ' + (($cl.Out + ' ' + $cl.Err).Trim() -replace '\s+', ' ').Substring(0, [math]::Min(300, ($cl.Out + ' ' + $cl.Err).Trim().Length))
            Show-HelperIdle; return
        }
        $answerLabel.Text = 'Сохраняю библиотеку на GitHub…'
        [void](Invoke-External 'git add -A' $lib 120)
        [void](Invoke-External ('git commit -m "Разбор новых постов канала ' + (Get-Date -Format 'dd.MM.yyyy HH:mm') + ' (Библиотекарь)"') $lib 120)
        $push = Invoke-External 'git push' $lib 180
        $pushNote = if ($push.Code -eq 0) { 'Сохранено на GitHub.' } else { 'На GitHub отправить не удалось — сохранено на компьютере.' }
        $answerLabel.Text = $summary.Trim() + ' ' + $pushNote + ' Нажмите «Сайт промптов», чтобы посмотреть.'
        Show-Helper 'found'
    } finally { $promptsButton.Enabled = $true }
}
# ---------------- Разбор картинок: битые и копии — в Корзину, непонятные названия — по смыслу ----------------
function Get-ImageKind([string]$Path) {
    # By content, not by extension: a «.png» that holds an HTML/XML error page is broken.
    $b = [byte[]]::new(16)
    try { $s = [IO.File]::OpenRead($Path); $n = $s.Read($b, 0, 16); $s.Close() } catch { return 'unreadable' }
    if ($n -lt 4) { return 'empty' }
    if ($b[0] -eq 0x89 -and $b[1] -eq 0x50) { return 'png' }
    if ($b[0] -eq 0xFF -and $b[1] -eq 0xD8) { return 'jpg' }
    if ($b[0] -eq 0x47 -and $b[1] -eq 0x49) { return 'gif' }
    if ($b[0] -eq 0x42 -and $b[1] -eq 0x4D) { return 'bmp' }
    if ($b[0] -eq 0x52 -and $b[1] -eq 0x49 -and $b[8] -eq 0x57) { return 'webp' }
    if ($n -ge 12 -and [Text.Encoding]::ASCII.GetString($b, 4, 4) -eq 'ftyp') { return 'heif' }
    return 'notimage'
}
function Get-ImageDHash([string]$Path) {
    # 64-bit «difference hash» of a 9x8 grey thumbnail: near-identical pictures differ in few bits.
    $img = $null; $bmp = $null
    try {
        $img = [System.Drawing.Image]::FromFile($Path)
        $bmp = New-Object System.Drawing.Bitmap 9, 8
        $g = [System.Drawing.Graphics]::FromImage($bmp); $g.InterpolationMode = 'HighQualityBilinear'; $g.DrawImage($img, 0, 0, 9, 8); $g.Dispose()
        $bits = [uint64]0
        for ($y = 0; $y -lt 8; $y++) { for ($x = 0; $x -lt 8; $x++) {
            $a = $bmp.GetPixel($x, $y); $c = $bmp.GetPixel($x + 1, $y)
            $bits = ($bits -shl 1) -bor [uint64]$(if ((0.299*$a.R + 0.587*$a.G + 0.114*$a.B) -gt (0.299*$c.R + 0.587*$c.G + 0.114*$c.B)) { 1 } else { 0 })
        } }
        return [pscustomobject]@{ Hash = $bits; W = $img.Width; H = $img.Height }
    } catch { return $null } finally { if ($bmp) { $bmp.Dispose() }; if ($img) { $img.Dispose() } }
}
function Get-BitDistance([uint64]$A, [uint64]$B) { $x = $A -bxor $B; $n = 0; while ($x) { $n += [int]($x -band 1); $x = $x -shr 1 }; return $n }
function Test-GenericImageName([string]$BaseName) {
    $n = $BaseName.Trim()
    return ($n -match '(?i)^(image|img|photo|picture|pic|generated image|chatgpt image|grok-image|gpt-image|leonardo|krea|dall|midjourney|mj_|sora|untitled|screenshot|скриншот|снимок|изображение|file|download|unnamed|без названия|photo_\d)' -or
            $n -match '^[\d\s_().-]+$' -or $n -match '(?i)^[0-9a-f]{8}-[0-9a-f]{4}-' -or $n -match '(?i)^[0-9a-f]{16,}' -or $n -match '^[-_\s]{3,}' -or
            ($n -match '^[a-z0-9_\-]{25,}$' -and $n -notmatch '[aeiouy]{1}[a-z]{2,}_[a-z]{3,}_[a-z]{3,}'))
}
function Get-ImageCleanupPlan([string]$Folder) {
    $files = @(Get-ChildItem -LiteralPath $Folder -File -ErrorAction SilentlyContinue | Where-Object { $_.Extension -match '^\.(png|jpe?g|jfif|webp|gif|bmp|avif|heic)$' })
    $broken = New-Object System.Collections.Generic.List[object]
    $ok = New-Object System.Collections.Generic.List[object]
    $i = 0
    foreach ($f in $files) {
        $i++
        if (($i % 15) -eq 0) { $answerLabel.Text = "Смотрю картинки: $i из $($files.Count)…"; [Windows.Forms.Application]::DoEvents() }
        $kind = Get-ImageKind $f.FullName
        if ($kind -in 'notimage','empty','unreadable') { $broken.Add($f); continue }
        $d = $null
        if ($kind -in 'png','jpg','gif','bmp') { $d = Get-ImageDHash $f.FullName; if (!$d) { $broken.Add($f); continue } }
        $sha = try { (Get-FileHash -LiteralPath $f.FullName -Algorithm SHA256).Hash } catch { '' }
        $ok.Add([pscustomobject]@{ F = $f; Sha = $sha; D = $d; Generic = (Test-GenericImageName $f.BaseName) })
    }
    # Which copy stays: bigger picture, then the one with a meaningful name, then bigger file.
    $rank = { param($x) $px = if ($x.D) { [long]$x.D.W * $x.D.H } else { 0 }; @($px, $(if ($x.Generic) { 0 } else { 1 }), [long]$x.F.Length) }
    $better = { param($a, $b) $ra = & $rank $a; $rb = & $rank $b; for ($k = 0; $k -lt 3; $k++) { if ($ra[$k] -ne $rb[$k]) { return ($ra[$k] -gt $rb[$k]) } }; return ($a.F.Name.Length -le $b.F.Name.Length) }
    $exact = New-Object System.Collections.Generic.List[object]
    $dropped = @{}
    foreach ($g in @($ok | Where-Object { $_.Sha } | Group-Object Sha | Where-Object Count -gt 1)) {
        $keep = $g.Group[0]; foreach ($x in $g.Group) { if (& $better $x $keep) { $keep = $x } }
        foreach ($x in $g.Group) { if ($x -ne $keep) { $exact.Add([pscustomobject]@{ Keep = $keep; Drop = $x }); $dropped[$x.F.FullName] = $true } }
    }
    # Similar-looking pairs: same proportions and at most 4 different bits of 64.
    $answerLabel.Text = 'Сравниваю картинки между собой…'; [Windows.Forms.Application]::DoEvents()
    $cand = @($ok | Where-Object { $_.D -and !$dropped.ContainsKey($_.F.FullName) })
    $similar = New-Object System.Collections.Generic.List[object]
    for ($a = 0; $a -lt $cand.Count; $a++) {
        if (($a % 40) -eq 0) { [Windows.Forms.Application]::DoEvents() }
        for ($b = $a + 1; $b -lt $cand.Count; $b++) {
            $x = $cand[$a]; $y = $cand[$b]
            if ([math]::Abs(($x.D.W / $x.D.H) - ($y.D.W / $y.D.H)) -gt 0.02) { continue }
            $dist = Get-BitDistance $x.D.Hash $y.D.Hash
            if ($dist -gt 4) { continue }
            $keep = if (& $better $x $y) { $x } else { $y }; $drop = if ($keep -eq $x) { $y } else { $x }
            $similar.Add([pscustomobject]@{ Keep = $keep; Drop = $drop; Dist = $dist; Sure = ($dist -le 1 -and $x.D.W -eq $y.D.W -and $x.D.H -eq $y.D.H) })
        }
    }
    $generic = @($ok | Where-Object { $_.Generic -and !$dropped.ContainsKey($_.F.FullName) } | ForEach-Object { $_.F.FullName })
    return [pscustomobject]@{ Folder = $Folder; Total = $files.Count; Broken = $broken; Exact = $exact; Similar = $similar; Generic = $generic }
}
function New-PairThumb([string]$Left, [string]$Right) {
    $bmp = New-Object System.Drawing.Bitmap 256, 128
    $g = [System.Drawing.Graphics]::FromImage($bmp); $g.Clear([System.Drawing.Color]::White); $g.InterpolationMode = 'HighQualityBicubic'
    $i = 0
    foreach ($p in @($Left, $Right)) {
        try {
            $img = [System.Drawing.Image]::FromFile($p)
            $scale = [math]::Min(124 / $img.Width, 124 / $img.Height)
            $w = [int]($img.Width * $scale); $h = [int]($img.Height * $scale)
            $g.DrawImage($img, [int]($i * 128 + (128 - $w) / 2), [int]((128 - $h) / 2), $w, $h); $img.Dispose()
        } catch { }
        $i++
    }
    $pen = New-Object System.Drawing.Pen ([System.Drawing.Color]::FromArgb(200,60,40)), 3
    $g.DrawRectangle($pen, 129, 1, 125, 125); $pen.Dispose(); $g.Dispose()
    return $bmp
}
function Show-ImageCleanupDialog($Plan) {
    $dlg = New-Object System.Windows.Forms.Form
    $dlg.Text = 'Разбор картинок: ' + $Plan.Folder; $dlg.StartPosition = 'CenterParent'; $dlg.Font = $form.Font
    $dlg.Size = [System.Drawing.Size]::new(900,700); $dlg.MinimumSize = [System.Drawing.Size]::new(700,500)
    $info = New-Object System.Windows.Forms.Label; $info.Dock = 'Top'; $info.Height = 120; $info.Padding = [System.Windows.Forms.Padding]::new(10,8,10,0)
    $bytes = [long]((@($Plan.Broken | ForEach-Object { $_.Length }) + @($Plan.Exact | ForEach-Object { $_.Drop.F.Length })) | Measure-Object -Sum).Sum
    $info.Text = "Картинок в папке: $($Plan.Total).`r`n" +
        "• Битых (не открываются — вместо картинки ошибка скачивания): $($Plan.Broken.Count) — в Корзину.`r`n" +
        "• Точных копий: $($Plan.Exact.Count) — в Корзину (остаётся одна, самая крупная).`r`n" +
        "• Похожих пар: $($Plan.Similar.Count) — ниже: СЛЕВА остаётся, СПРАВА (в красной рамке) уходит в Корзину. Отметьте пары, где картинки действительно одинаковые; заранее отмечены только почти точные совпадения.`r`n" +
        "Освободится сразу: $(Format-Bytes $bytes) + отмеченные пары."
    $list = New-Object System.Windows.Forms.ListView
    $list.Dock = 'Fill'; $list.View = 'LargeIcon'; $list.CheckBoxes = $true
    $imgs = New-Object System.Windows.Forms.ImageList; $imgs.ImageSize = [System.Drawing.Size]::new(256,128); $imgs.ColorDepth = 'Depth24Bit'
    $list.LargeImageList = $imgs
    $k = 0
    foreach ($p in $Plan.Similar) {
        $imgs.Images.Add((New-PairThumb $p.Keep.F.FullName $p.Drop.F.FullName))
        $it = New-Object System.Windows.Forms.ListViewItem ($(if ($p.Sure) { 'одинаковые' } else { 'похожи, разница ' + $p.Dist }) + ' → убрать: ' + $p.Drop.F.Name), $k
        $it.Checked = [bool]$p.Sure; $it.Tag = $p; [void]$list.Items.Add($it); $k++
        if (($k % 10) -eq 0) { [Windows.Forms.Application]::DoEvents() }
    }
    $bottomBar = New-Object System.Windows.Forms.Panel; $bottomBar.Dock = 'Bottom'; $bottomBar.Height = 92
    $renameCheck = New-Object System.Windows.Forms.CheckBox; $renameCheck.Location = [System.Drawing.Point]::new(10,8); $renameCheck.Size = [System.Drawing.Size]::new(860,40)
    $mins = [math]::Max(1, [math]::Ceiling($Plan.Generic.Count / 20 * 1.6))
    $renameCheck.Text = if (Get-Command claude.cmd -ErrorAction SilentlyContinue) {
        "Переименовать по смыслу картинки с непонятными названиями ($($Plan.Generic.Count) шт.): Claude посмотрит каждую. Примерно $mins мин., расходует лимит подписки Claude. Вернуть старые названия — «Отменить перенос»."
    } else {
        "Дать картинкам с непонятными названиями ($($Plan.Generic.Count) шт.) аккуратные имена по дате: «Картинка 2026-10-11 001.png». Вернуть старые названия — «Отменить перенос»."
    }
    $renameCheck.Checked = ($Plan.Generic.Count -gt 0); $renameCheck.Enabled = ($Plan.Generic.Count -gt 0)
    $all = New-Object System.Windows.Forms.Button; $all.Text = 'Отметить все пары'; $all.Location = [System.Drawing.Point]::new(10,52); $all.Size = [System.Drawing.Size]::new(150,30)
    $all.Add_Click({ foreach ($it in $list.Items) { $it.Checked = $true } })
    $ok = New-Object System.Windows.Forms.Button; $ok.Text = 'Выполнить'; $ok.Location = [System.Drawing.Point]::new(650,52); $ok.Size = [System.Drawing.Size]::new(110,30); $ok.DialogResult = 'OK'; $ok.Anchor = 'Top,Right'
    $cancel = New-Object System.Windows.Forms.Button; $cancel.Text = 'Отмена'; $cancel.Location = [System.Drawing.Point]::new(766,52); $cancel.Size = [System.Drawing.Size]::new(104,30); $cancel.DialogResult = 'Cancel'; $cancel.Anchor = 'Top,Right'
    $bottomBar.Controls.AddRange(@($renameCheck,$all,$ok,$cancel))
    $dlg.Controls.Add($list); $dlg.Controls.Add($info); $dlg.Controls.Add($bottomBar)
    $dlg.CancelButton = $cancel
    try {
        if ($dlg.ShowDialog($form) -ne 'OK') { return @{ Ok = $false } }
        $drop = New-Object System.Collections.Generic.List[object]
        foreach ($b in $Plan.Broken) { $drop.Add([pscustomobject]@{ Path = $b.FullName; Reason = 'битая картинка (ошибка скачивания)'; Bytes = $b.Length }) }
        foreach ($e in $Plan.Exact) { $drop.Add([pscustomobject]@{ Path = $e.Drop.F.FullName; Reason = 'точная копия ' + $e.Keep.F.Name; Bytes = $e.Drop.F.Length }) }
        $seen = @{}
        foreach ($it in $list.Items) {
            if (!$it.Checked) { continue }
            $p = $it.Tag; $dp = $p.Drop.F.FullName
            # A picture chosen to stay in one pair is never thrown away by another pair.
            if ($seen.ContainsKey($dp)) { continue }
            $keepIsDropped = @($drop | Where-Object { $_.Path -eq $p.Keep.F.FullName }).Count -gt 0
            if ($keepIsDropped) { continue }
            $drop.Add([pscustomobject]@{ Path = $dp; Reason = 'похожа на ' + $p.Keep.F.Name; Bytes = $p.Drop.F.Length }); $seen[$dp] = $true
        }
        return @{ Ok = $true; Drop = $drop; Rename = $renameCheck.Checked }
    } finally { $imgs.Dispose(); $dlg.Dispose() }
}
function Get-SafeImageName([string]$Text) {
    $invalid = [IO.Path]::GetInvalidFileNameChars()
    $s = (-join @($Text.Trim().ToCharArray() | ForEach-Object { if ($invalid -contains $_) { ' ' } else { [string]$_ } }))
    $s = ([regex]::Replace($s, '\s+', ' ')).Trim(' ', '.', ',')
    if ($s.Length -gt 80) { $s = $s.Substring(0, 80).Trim() }
    if ($s) { $s = $s.Substring(0,1).ToUpperInvariant() + $s.Substring(1) }
    return $s
}
function Invoke-ImageDateRename($Paths, [string]$LogPath) {
    # Without Claude Code: «Картинка 2026-10-11 001.png» by the date the picture was saved.
    $renamed = 0; $counters = @{}
    $todo = @($Paths | Where-Object { Test-Path -LiteralPath $_ -PathType Leaf } | ForEach-Object { Get-Item -LiteralPath $_ } | Sort-Object LastWriteTime)
    foreach ($f in $todo) {
        $day = $f.LastWriteTime.ToString('yyyy-MM-dd'); $key = $f.DirectoryName + '|' + $day
        $ext = $f.Extension.ToLowerInvariant()
        do {
            $counters[$key] = 1 + [int]$counters[$key]
            $target = Join-Path $f.DirectoryName ('Картинка {0} {1:000}{2}' -f $day, $counters[$key], $ext)
        } while (Test-Path -LiteralPath $target)
        try { Move-Item -LiteralPath $f.FullName -Destination $target -ErrorAction Stop; Add-MoveLogEntry $LogPath 'ПЕРЕНЕСЁН' $f.FullName $target 'название по дате'; $renamed++ } catch { }
    }
    return "Переименовано по дате: $renamed из $($todo.Count)."
}
function Invoke-ImageRename($Paths, [string]$LogPath) {
    $claude = (Get-Command claude.cmd -ErrorAction SilentlyContinue).Source
    if (!$claude) { return (Invoke-ImageDateRename $Paths $LogPath) }
    $guide = Join-Path (Split-Path $script:MediaDir -Parent) 'ИМЕНА_КАРТИНОК.md'
    $work = Join-Path $env:TEMP ('librarian_names_' + (Get-Date -Format 'HHmmss'))
    [void][IO.Directory]::CreateDirectory($work)
    Copy-Item -LiteralPath $guide -Destination (Join-Path $work 'ИМЕНА_КАРТИНОК.md')
    $todo = @($Paths | Where-Object { Test-Path -LiteralPath $_ })
    $renamed = 0; $batchSize = 20
    for ($start = 0; $start -lt $todo.Count; $start += $batchSize) {
        $batch = @($todo[$start..([math]::Min($start + $batchSize, $todo.Count) - 1)])
        $answerLabel.Text = "Даю картинкам понятные названия: $start из $($todo.Count)… (можно заниматься своими делами)"
        ConvertTo-Json -InputObject $batch | Set-Content -LiteralPath (Join-Path $work 'batch.json') -Encoding UTF8
        Remove-Item -LiteralPath (Join-Path $work 'names.json') -ErrorAction SilentlyContinue
        $dirs = @($batch | ForEach-Object { [IO.Path]::GetDirectoryName($_) } | Select-Object -Unique)
        $addDirs = ($dirs | ForEach-Object { '--add-dir "' + $_ + '"' }) -join ' '
        $ask = 'Сделай задание из файла ИМЕНА_КАРТИНОК.md в этой папке.'
        [void](Invoke-External ('"' + $claude + '" -p "' + $ask + '" ' + $addDirs + ' --allowedTools "Read,Write"') $work 900)
        $namesFile = Join-Path $work 'names.json'
        if (!(Test-Path -LiteralPath $namesFile)) { continue }
        try { $names = Get-Content -LiteralPath $namesFile -Raw -Encoding UTF8 | ConvertFrom-Json } catch { continue }
        foreach ($prop in $names.PSObject.Properties) {
            $src = [string]$prop.Name
            if (!(Test-Path -LiteralPath $src -PathType Leaf) -or $batch -notcontains $src) { continue }
            $new = Get-SafeImageName ([string]$prop.Value)
            if (!$new) { continue }
            $dir = [IO.Path]::GetDirectoryName($src); $ext = [IO.Path]::GetExtension($src).ToLowerInvariant()
            $target = Join-Path $dir ($new + $ext); $n = 2
            while (Test-Path -LiteralPath $target) { $target = Join-Path $dir ("{0} ({1}){2}" -f $new, $n, $ext); $n++ }
            try { Move-Item -LiteralPath $src -Destination $target -ErrorAction Stop; Add-MoveLogEntry $LogPath 'ПЕРЕНЕСЁН' $src $target 'новое название'; $renamed++ } catch { }
        }
    }
    Remove-Item -LiteralPath $work -Recurse -Force -ErrorAction SilentlyContinue
    return "Переименовано: $renamed из $($todo.Count)."
}
function Invoke-ImageCleanup([string]$Folder) {
    if (!$Folder -or !(Test-Path -LiteralPath $Folder -PathType Container)) { $answerLabel.Text = 'Не нашёл папку с картинками.'; $script:TaskDone = $false; return }
    Start-HelperLoop 'search'
    $plan = Get-ImageCleanupPlan $Folder
    if ($plan.Total -eq 0) { $answerLabel.Text = 'В этой папке нет картинок: ' + $Folder; Show-HelperIdle; return }
    $answerLabel.Text = "Нашёл: битых $($plan.Broken.Count), точных копий $($plan.Exact.Count), похожих пар $($plan.Similar.Count), с непонятными названиями $($plan.Generic.Count). Проверьте и нажмите «Выполнить»."
    Show-HelperIdle
    $choice = Show-ImageCleanupDialog $plan
    if (!$choice.Ok) { $answerLabel.Text = 'Хорошо, картинки не трогаю.'; return }
    [void][IO.Directory]::CreateDirectory($script:MoveLogDir)
    $logPath = Join-Path $script:MoveLogDir ('moves_' + (Get-Date -Format 'yyyyMMdd_HHmmss') + '_картинки.csv')
    Show-Helper 'sweep'
    $binned = 0; $freed = [long]0
    foreach ($d in $choice.Drop) {
        try {
            if (!(Test-Path -LiteralPath $d.Path -PathType Leaf)) { continue }
            [Microsoft.VisualBasic.FileIO.FileSystem]::DeleteFile($d.Path, [Microsoft.VisualBasic.FileIO.UIOption]::OnlyErrorDialogs, [Microsoft.VisualBasic.FileIO.RecycleOption]::SendToRecycleBin, [Microsoft.VisualBasic.FileIO.UICancelOption]::ThrowException)
            Add-MoveLogEntry $logPath 'В КОРЗИНЕ' $d.Path '' $d.Reason; $binned++; $freed += [long]$d.Bytes
        } catch { Add-MoveLogEntry $logPath 'ОШИБКА' $d.Path '' $_.Exception.Message }
    }
    $renameNote = ''
    if ($choice.Rename) {
        Start-HelperLoop 'search' -Silent
        $dropSet = @{}; foreach ($d in $choice.Drop) { $dropSet[$d.Path] = $true }
        $renameNote = ' ' + (Invoke-ImageRename @($plan.Generic | Where-Object { !$dropSet.ContainsKey($_) }) $logPath)
    }
    Update-Catalog
    $answerLabel.Text = "Готово: в Корзину $binned (освобождено $(Format-Bytes $freed)).$renameNote Вернуть названия — «Отменить перенос», удалённое — в Корзине."
    Show-Helper 'dust'
}
# ---------------- Новая папка: «создай папку для договоров с Ивановым» ----------------
function Get-FolderSuggestion([string]$Text, [string]$Storage) {
    # Same topics as the sorting plan; only shelves that really exist are offered.
    $t = $Text.ToLowerInvariant().Replace('ё','е')
    $rules = @(
        @('клиент', '01_РАБОТА\КЛИЕНТЫ'),
        @('копирайт|тексты для', '01_РАБОТА\КОПИРАЙТИНГ'),
        @('(^|[^a-z])hr([^a-z]|$)|ваканс|кандидат|резюме|подбор|кадр|собеседов', '01_РАБОТА\HR'),
        @('самозанят|налог|выплат', '01_РАБОТА\САМОЗАНЯТОСТЬ'),
        @('работ|заказ|договор|сч[её]т|акт[ыа ]|смет', '01_РАБОТА'),
        @('портфолио|мои работы|готов\S* работ|кейс|конкурс|мультфильм|мультик', '03_МОИ_РАБОТЫ'),
        @('проект|бот|сайт|приложени', '02_МОИ_ПРОЕКТЫ'),
        @('семь|мам|пап[аы]|дет[ейи]|ребен|личн|дом[аш ]|здоров|врач|медиц|документ|паспорт|квартир|машин|отпуск|путешеств', '04_ЛИЧНОЕ_И_СЕМЬЯ'),
        @('обучен|курс|урок|учеб|лекци|книг|вебинар|психолог', '05_ОБУЧЕНИЕ'))
    # Own projects and rules from «настройки.json» come first.
    $all = New-Object System.Collections.Generic.List[object]
    foreach ($p in $script:Projects) { $all.Add(@([string]$p.'Слова', ('02_МОИ_ПРОЕКТЫ\' + $p.'Папка'))) }
    foreach ($x in $script:ExtraRules) { $all.Add(@([string]$x.'Слова', ([string]$x.'Полка').Replace('/', '\').Trim('\'))) }
    foreach ($r in $rules) { $all.Add($r) }
    foreach ($r in $all) {
        if ($t -match $r[0]) {
            $rel = $r[1]
            while ($rel -and !(Test-Path -LiteralPath (Join-Path $Storage $rel) -PathType Container)) { $rel = Split-Path $rel -Parent }
            if ($rel) { return $rel }
        }
    }
    return '90_РАЗОБРАТЬ'
}
function Get-FolderNameFromSpeech([string]$Text) {
    # «её зовут Пекарня», «назови Пекарня», «называется …», «пусть будет …», «имя …»
    $m = [regex]::Match($Text, '(?i)(?:зовут|называ\w*|назов\w*|назвать|назван\w*|пусть\s+будет|имя)\s+(?:(?:её|ее|его|она|папку|папка|как)\s+)?[«"]?([\p{L}\p{Nd}][\p{L}\p{Nd}\- ]{0,40}?)[»"]?(?=\s*[,.;!?]|\s+(?:а|и|для|там|будет|будем|чтобы|она|мы)\b|\s*$)')
    if (!$m.Success) { return '' }
    $n = $m.Groups[1].Value.Trim()
    if (!$n) { return '' }
    return $n.Substring(0,1).ToUpperInvariant() + $n.Substring(1)
}
function Resolve-FolderName([string]$Raw, [string]$Purpose) {
    # Dictation may keep typing into the name box: drop the words of the command itself.
    $spoken = Get-FolderNameFromSpeech ($Raw + ' . ' + $Purpose)
    $junk = @('а','ну','вот','на','надо','нужно','нужна','создать','создай','создайте','сделай','новую','новая','новой','папку','папка','папки','назови','назовите','назвать','её','ее','его','пусть','будет','называется','название','зовут','это','мне','пожалуйста','как','и')
    $words = @($Raw.Trim().TrimEnd('.', '!', '?', ',') -split '\s+' | Where-Object { $_ -and ($junk -notcontains $_.ToLowerInvariant().Trim(',', '.')) })
    $clean = ($words -join ' ').Trim(' ', ',', '.', '«', '»', '"')
    if ($spoken -and ($Raw -match '(?i)созда|папк|зовут|назов|назва|пусть\s+будет' -or !$clean)) { $clean = $spoken }
    if ($clean) { $clean = $clean.Substring(0,1).ToUpperInvariant() + $clean.Substring(1) }
    $invalid = [IO.Path]::GetInvalidFileNameChars()
    return (-join @($clean.ToCharArray() | ForEach-Object { if ($invalid -contains $_) { '_' } else { [string]$_ } })).TrimEnd('.', ' ')
}
function Invoke-CreateFolder([string]$Text) {
    $storage = Read-SmallTextFile $script:StorageRootConfigPath
    if (!$storage -or !(Test-Path -LiteralPath $storage -PathType Container)) { $storage = Get-StorageRoot; if (!$storage) { return } }
    # What was already said: «папку <имя> для <назначение>».
    $t = $Text.Trim().TrimEnd('.', '!', '?')
    $name = ''; $purpose = ''
    $m = [regex]::Match($t, '(?i)папк\S*\s+(.*)$')
    if ($m.Success) {
        $rest = $m.Groups[1].Value
        $p = [regex]::Match($rest, '(?i)^(.*?)\s*\b(для|под|про|чтобы)\b\s+(.*)$')
        if ($p.Success) { $name = $p.Groups[1].Value; $purpose = $p.Groups[3].Value } else { $name = $rest }
        $name = ([regex]::Replace($name, '(?i)^(новую|новая|новой|одну|еще|ещё|мне|пожалуйста|с\s+названием|под\s+названием|называется)\s+', '')).Trim(' ', '«', '»', '"')
        if ($name -match '(?i)^(новую|новая|мне|пожалуйста)$') { $name = '' }
    }
    $spoken = Get-FolderNameFromSpeech $t
    if ($spoken) { $name = $spoken } elseif ($name) { $name = Resolve-FolderName $name '' }
    # Folders of the shelves, two levels deep, for the «where» list.
    $places = New-Object System.Collections.Generic.List[string]
    foreach ($shelf in @(Get-ChildItem -LiteralPath $storage -Directory -ErrorAction SilentlyContinue | Sort-Object Name)) {
        $places.Add($shelf.Name)
        foreach ($sub in @(Get-ChildItem -LiteralPath $shelf.FullName -Directory -ErrorAction SilentlyContinue | Sort-Object Name)) { $places.Add($shelf.Name + '\' + $sub.Name) }
    }
    $dlg = New-Object System.Windows.Forms.Form
    $dlg.Text = 'Новая папка'; $dlg.StartPosition = 'CenterParent'; $dlg.Font = $form.Font
    $dlg.Size = [System.Drawing.Size]::new(640,330); $dlg.FormBorderStyle = 'FixedDialog'; $dlg.MinimizeBox = $false; $dlg.MaximizeBox = $false
    $l1 = New-Object System.Windows.Forms.Label; $l1.Text = 'Как назвать папку?'; $l1.Location = [System.Drawing.Point]::new(12,12); $l1.AutoSize = $true
    $nameBox = New-Object System.Windows.Forms.TextBox; $nameBox.Location = [System.Drawing.Point]::new(12,32); $nameBox.Size = [System.Drawing.Size]::new(600,26); $nameBox.Text = $name
    $l2 = New-Object System.Windows.Forms.Label; $l2.Text = 'Для чего она? (по этому я выберу полку)'; $l2.Location = [System.Drawing.Point]::new(12,66); $l2.AutoSize = $true
    $purposeBox = New-Object System.Windows.Forms.TextBox; $purposeBox.Location = [System.Drawing.Point]::new(12,86); $purposeBox.Size = [System.Drawing.Size]::new(600,26); $purposeBox.Text = $purpose
    $l3 = New-Object System.Windows.Forms.Label; $l3.Text = 'Куда поставить:'; $l3.Location = [System.Drawing.Point]::new(12,120); $l3.AutoSize = $true
    $placeBox = New-Object System.Windows.Forms.ComboBox; $placeBox.DropDownStyle = 'DropDownList'; $placeBox.Location = [System.Drawing.Point]::new(12,140); $placeBox.Size = [System.Drawing.Size]::new(600,26)
    foreach ($pl in $places) { [void]$placeBox.Items.Add($pl) }
    $preview = New-Object System.Windows.Forms.Label; $preview.Location = [System.Drawing.Point]::new(12,176); $preview.Size = [System.Drawing.Size]::new(600,54); $preview.ForeColor = [System.Drawing.Color]::FromArgb(76,59,37)
    $state = @{ Manual = $false; Busy = $false }
    $update = {
        if (!$state.Manual) {
            $sug = Get-FolderSuggestion ($nameBox.Text + ' ' + $purposeBox.Text) $storage
            $i = $placeBox.Items.IndexOf($sug)
            if ($i -lt 0) { $i = $placeBox.Items.IndexOf('90_РАЗОБРАТЬ') }
            $state.Busy = $true; if ($i -ge 0) { $placeBox.SelectedIndex = $i }; $state.Busy = $false
        }
        $nm = Resolve-FolderName $nameBox.Text $purposeBox.Text
        $preview.Text = if ($nm) { 'Будет создана папка «' + $nm + '»:  ' + (Join-Path (Join-Path $storage ([string]$placeBox.SelectedItem)) $nm) } else { 'Скажите или напишите название папки (например: «Пекарня»).' }
    }
    $nameBox.Add_TextChanged($update); $purposeBox.Add_TextChanged($update)
    $placeBox.Add_SelectedIndexChanged({ if (!$state.Busy) { $state.Manual = $true }; $nm = Resolve-FolderName $nameBox.Text $purposeBox.Text; $preview.Text = if ($nm) { 'Будет создана папка «' + $nm + '»:  ' + (Join-Path (Join-Path $storage ([string]$placeBox.SelectedItem)) $nm) } else { 'Скажите или напишите название папки (например: «Пекарня»).' } })
    $ok = New-Object System.Windows.Forms.Button; $ok.Text = 'Создать'; $ok.Location = [System.Drawing.Point]::new(392,240); $ok.Size = [System.Drawing.Size]::new(110,32); $ok.DialogResult = 'OK'
    $cancel = New-Object System.Windows.Forms.Button; $cancel.Text = 'Отмена'; $cancel.Location = [System.Drawing.Point]::new(508,240); $cancel.Size = [System.Drawing.Size]::new(104,32); $cancel.DialogResult = 'Cancel'
    $dlg.Controls.AddRange(@($l1,$nameBox,$l2,$purposeBox,$l3,$placeBox,$preview,$ok,$cancel)); $dlg.AcceptButton = $ok; $dlg.CancelButton = $cancel
    & $update
    $dlg.Add_Shown({ if ($nameBox.Text) { [void]$purposeBox.Focus() } else { [void]$nameBox.Focus() } })
    Start-HelperLoop 'write' -Silent   # he sits down to write while you fill in the name
    try {
        if ($dlg.ShowDialog($form) -ne 'OK') { $answerLabel.Text = 'Хорошо, папку не создаю.'; Show-HelperIdle; return }
        $folderName = Resolve-FolderName $nameBox.Text $purposeBox.Text
        if (!$folderName) { $answerLabel.Text = 'Не понял, как назвать папку. Скажите ещё раз, например: «создай папку для нового клиента, её зовут Пекарня».'; $script:TaskDone = $false; return }
        $where = [string]$placeBox.SelectedItem
        $full = Join-Path (Join-Path $storage $where) $folderName
        if (Test-Path -LiteralPath $full) {
            $answerLabel.Text = 'Такая папка уже есть — открыл её: ' + $full
        } else {
            [void][IO.Directory]::CreateDirectory($full)
            $note = $purposeBox.Text.Trim()
            if ($note) { Set-Content -LiteralPath (Join-Path $full '_для_чего_эта_папка.txt') -Value ("Создана " + (Get-Date -Format 'dd.MM.yyyy') + ".`r`nДля чего: " + $note) -Encoding UTF8 }
            $nice = (@($where -split '\\') | ForEach-Object { $s = ([regex]::Replace($_, '^\d+_', '')).Replace('_',' ').ToLowerInvariant(); if ($s) { $s.Substring(0,1).ToUpperInvariant() + $s.Substring(1) } }) -join ' → '
            $answerLabel.Text = 'Создал папку «' + $folderName + '» на полке «' + $nice + '». Путь скопирован (Ctrl+V, чтобы вставить): ' + $full
            Show-Helper 'write'   # «Так, запишем… Новая папка готова, ставлю её на полочку!»
        }
        try { [Windows.Forms.Clipboard]::SetText($full) } catch { }
        $selectedPathBox.Text = $full
        Start-Process explorer.exe -ArgumentList "/select,`"$full`""
    } finally { $dlg.Dispose() }
}
function Set-RequestMode([string]$Mode) {
    # «Поиск» only ever searches; «Задание» carries out what is asked. The box turns orange in task mode.
    $isTask = ($Mode -eq 'task')
    if ($modeTask.Checked -ne $isTask) { $modeTask.Checked = $isTask; $modeFind.Checked = !$isTask }
    if ($isTask) {
        $searchBox.BackColor = [System.Drawing.Color]::FromArgb(255,226,190)
        $searchLabel.Text = 'ЗАДАНИЕ: скажите или напишите, что сделать — «удали папку вакансии», «эту папку удали», «открой папку претензия», «сделай уборку»:'
        $answerLabel.Text = 'Режим «Задание». Я выполню то, что вы скажете; перед удалением всегда спрошу. После — вернусь в «Поиск».'
    } else {
        $searchBox.BackColor = [System.Drawing.SystemColors]::Window
        $searchLabel.Text = 'Что найти? Напишите или нажмите «Сказать» — например: «договор с Ивановым», «фото собаки», «презентация про отпуск»:'
    }
}
function Invoke-Request([switch]$FromVoice) {
    $text = $searchBox.Text.Trim()
    if (!$modeTask.Checked) { Apply-Search -Speak:(!$FromVoice); return }
    if (!$text) { return }
    $script:TaskDone = $true
    Invoke-Task $text
    if ($script:TaskDone) { $modeFind.Checked = $true; Set-RequestMode 'find' }
}
function Invoke-Task([string]$text) {
    # Sets $script:TaskDone = $false when the task was not understood, so task mode stays on for another try.
    $t = $text.ToLowerInvariant()
    # «разбери картинки в загрузках / в этой папке / в папке X» — broken, copies, look-alikes, names.
    if ($t -match '(разбер|разобра|почист|убер|навед|перебер|переимен|просмотр|посмотр)\S*.*(картин|фото|изображ)') {
        $p = Get-PathFromText $text
        $folder = $null
        if ($p -and !$p.StartsWith('?')) { $folder = if (Test-Path -LiteralPath $p -PathType Container) { $p } else { [IO.Path]::GetDirectoryName($p) } }
        elseif ($t -match 'загрузк') { $folder = Get-DownloadsPath }
        elseif ($t -match 'рабоч\S*\s+стол') { $folder = [Environment]::GetFolderPath('Desktop') }
        elseif ($t -match '(эт\S+|той)\s+папк|здесь|тут') {
            $sel = @($script:UserPick | Where-Object { $_ -and (Test-Path -LiteralPath $_) }); if ($sel.Count -eq 0) { $sel = @(Get-SelectedPaths) }
            if ($sel.Count) { $folder = [IO.Path]::GetDirectoryName($sel[0]) }
        } else {
            $m = [regex]::Match($t, 'папк\S*\s+(.+)$')
            if ($m.Success) { $found = @(Find-ByName $m.Groups[1].Value -Folders); if ($found.Count -eq 1) { $folder = $found[0] } elseif ($found.Count -gt 1) { $folder = Show-PickList 'Какую папку?' ($found | Select-Object -First 50) } }
        }
        if (!$folder) {
            $fb = New-Object System.Windows.Forms.FolderBrowserDialog; $fb.Description = 'В какой папке разобрать картинки?'
            if ($fb.ShowDialog() -eq 'OK') { $folder = $fb.SelectedPath } else { $answerLabel.Text = 'Хорошо, картинки не трогаю.'; return }
        }
        Invoke-ImageCleanup $folder
        return
    }
    $path = Get-PathFromText $text
    if ($path) {
        if ($path.StartsWith('?')) { $answerLabel.Text = 'Не нашёл такой папки или файла: ' + $path.Substring(1) + '. Проверьте путь: в Проводнике щёлкните по адресной строке и скопируйте его целиком.'; $script:TaskDone = $false; return }
        $intent = if ($t -match 'удал|корзин|выбрось|выброси|выкин|убери|снеси|сотри|стереть|не нужн') { 'delete' } elseif ($t -match 'откр|покаж|где\s') { 'open' } else { 'ask' }
        Invoke-PathTask $path $intent
        return
    }
    # «создай новую папку для договоров» — asks name and purpose, picks the shelf, creates it and gives the path.
    if ($t -match '(созда|сдела|завед|завест|нужн\S*\s+нов)\S*\s+(.*\s)?папк') { Invoke-CreateFolder $text; return }
    # «удали папку Telegram Desktop», «открой папку Претензия», «удали файл …» — found by name, no path needed.
    $task = Get-NameTask $text
    if ($task) {
        $intent = if ($task.Verb -match 'откр|покажи') { 'open' } else { 'delete' }
        $isFolder = ($task.Kind -eq 'папк')
        $what = $task.What
        $kind = if ($isFolder) { 'папку' } else { 'файл' }
        if (@((Get-SearchTerms $what).Stems).Count -eq 0) {
            # «эту папку можно удалить» — no name: the folder of the file selected in the list (or that file).
            # The row the user clicked themselves — not the one the search highlighted on its own.
            $sel = @($script:UserPick | Where-Object { $_ -and (Test-Path -LiteralPath $_) })
            if ($sel.Count -eq 0) { $sel = @(Get-SelectedPaths) }
            if ($sel.Count -eq 0) { $answerLabel.Text = "Какую $kind? Назовите её (например, «удали папку вакансии») или выделите файл в списке и скажите «эту папку удали»."; $script:TaskDone = $false; return }
            $pick = if ($isFolder) { [IO.Path]::GetDirectoryName($sel[0]) } else { $sel[0] }
            Invoke-PathTask $pick $intent
            return
        }
        $found = @(Find-ByName $what -Folders:$isFolder)
        if ($found.Count -eq 0) { $answerLabel.Text = "Не нашёл $kind «$what». Попробуйте другое название или перетащите её мышкой в это окно."; $script:TaskDone = $false; return }
        $pick = if ($found.Count -eq 1) { $found[0] } else { Show-PickList ("Какую $kind?") ($found | Select-Object -First 50) }
        if (!$pick) { $answerLabel.Text = 'Хорошо, ничего не делаю.'; return }
        Invoke-PathTask $pick $intent
        return
    }
    if ($t -match '(уборк|наведи\s+порядок|навести\s+порядок)') { Invoke-WeeklyCleanup; return }
    if ($script:HasPrompts -and $t -match 'сайт') { Open-PromptSite; return }
    if ($script:HasPrompts -and $t -match 'телеграм|канал|промпт|промт' -and $t -match 'собер|собрат|сбор|забер|забрат|нов|разбер|разобрат|обнов') { Invoke-PromptCollect; return }
    if ($t -match 'обнови\s+каталог|обновить\s+каталог') { Update-Catalog; $answerLabel.Text = "Каталог обновлён: файлов $($script:AllFiles.Count)."; return }
    if ($t -match 'отмени\s+перенос|верни\s+(файлы|обратно)') { Undo-LastMove; return }
    # «найди…», «где лежит…», «покажи…» said in task mode is simply a search: switch and search.
    if ($t -match '(^|\s)(найд|найт|поищ|покаж|где\s|ищу|отыщ)') {
        $modeFind.Checked = $true; Set-RequestMode 'find'
        Apply-Search -Speak
        $answerLabel.Text = 'Это поиск — переключился на «Поиск». ' + $answerLabel.Text
        return
    }
    $answerLabel.Text = 'Не понял задание. Скажите, например: «удали папку вакансии», «эту папку удали» (после щелчка по файлу), «открой папку претензия», «сделай уборку», «обнови каталог». Или переключитесь на «Поиск».'
    $script:TaskDone = $false
}
$searchButton.Add_Click({ $script:SearchTimer.Stop(); Invoke-Request })
$modeTask.Add_CheckedChanged({ if ($modeTask.Checked) { Set-RequestMode 'task'; [void]$searchBox.Focus(); $searchBox.SelectAll() } else { Set-RequestMode 'find' } })
# A folder or file dragged from Explorer onto the window: ask what to do with it.
foreach ($c in @($form, $searchBox, $grid, $answerLabel, $top, $bottom, $pathPanel)) {
    $c.AllowDrop = $true
    $c.Add_DragEnter({ if ($_.Data.GetDataPresent([Windows.Forms.DataFormats]::FileDrop)) { $_.Effect = [Windows.Forms.DragDropEffects]::Copy } })
    $c.Add_DragDrop({
        $dropped = @($_.Data.GetData([Windows.Forms.DataFormats]::FileDrop))
        if ($dropped.Count -eq 0) { return }
        $form.Activate()
        if ($dropped.Count -gt 1) { $answerLabel.Text = "Перетащено $($dropped.Count) — беру первое: $($dropped[0]). Остальные перетащите по одному." }
        Invoke-PathTask ([string]$dropped[0]) 'ask'
    })
}
$searchBox.Add_KeyDown({
    if ($_.KeyCode -eq [System.Windows.Forms.Keys]::Enter) {
        $script:SearchTimer.Stop()
        Invoke-Request
        $_.SuppressKeyPress = $true
        $_.Handled = $true
    }
})
# Search by itself a moment after typing or dictation stops. In task mode the task waits for Enter,
# unless it was dictated: then it starts by itself (deleting still asks for confirmation).
$script:SearchTimer = New-Object System.Windows.Forms.Timer
$script:SearchTimer.Interval = 1100
$script:SearchTimer.Add_Tick({
    $script:SearchTimer.Stop()
    $text = $searchBox.Text.Trim()
    if ($text.Length -lt 2) { return }
    if ($modeTask.Checked) {
        # Dictation ends a phrase with a full stop: such a task starts by itself (deleting still asks).
        if ($script:SpeakNext -or $text -match '[.!]$') { $script:SpeakNext = $false; Invoke-Request -FromVoice }
        else { $answerLabel.Text = 'Задание записано. Нажмите Enter — и я его выполню (перед удалением всё равно спрошу).' }
        return
    }
    Apply-Search
})
$searchBox.Add_TextChanged({ $script:SearchTimer.Stop(); $script:SearchTimer.Start() })
Add-Type -Namespace LibrarianApp -Name Keyboard -MemberDefinition '[DllImport("user32.dll")] public static extern void keybd_event(byte bVk, byte bScan, uint dwFlags, UIntPtr dwExtraInfo);'
$micButton.Add_Click({
    # Windows voice typing (Win+H) writes straight into the focused search box.
    # Arthur speaks before the microphone opens, so dictation does not type his words.
    if ($voiceCheck.Checked) { Play-Arthur 'listen' -Wait }
    [void]$searchBox.Focus(); $searchBox.SelectAll()
    $script:SpeakNext = $true
    [LibrarianApp.Keyboard]::keybd_event(0x5B, 0, 0, [UIntPtr]::Zero)
    [LibrarianApp.Keyboard]::keybd_event(0x48, 0, 0, [UIntPtr]::Zero)
    [LibrarianApp.Keyboard]::keybd_event(0x48, 0, 2, [UIntPtr]::Zero)
    [LibrarianApp.Keyboard]::keybd_event(0x5B, 0, 2, [UIntPtr]::Zero)
    $status.Text = 'Говорите, что найти. Когда закончите, поиск начнётся сам.'
})
$grid.Add_MouseUp({ if ($grid.SelectedItems.Count -gt 0) { $script:UserPick = [string]$grid.SelectedItems[0].Tag } })
$grid.Add_KeyUp({ if ($grid.SelectedItems.Count -gt 0) { $script:UserPick = [string]$grid.SelectedItems[0].Tag } })
$grid.Add_SelectedIndexChanged({
    if ($grid.SelectedItems.Count -gt 0) { $selectedPathBox.Text = [string]$grid.SelectedItems[0].Tag }
    else { $selectedPathBox.Clear() }
})
$copyPathButton.Add_Click({
    if ($selectedPathBox.Text) { try { [Windows.Forms.Clipboard]::SetText($selectedPathBox.Text) } catch { [void][LibrarianApp.LibBox]::Show('Не удалось скопировать путь.') } }
})
$grid.Add_ItemActivate({ Open-SelectedFile })
$form.Add_Shown({ [void]$searchBox.Focus() })
# Started by the weekly task: open straight into the cleanup.
$form.Add_Shown({
    if ($script:WeeklyMode) { $script:WeeklyMode = $false; Invoke-WeeklyCleanup; return }
    if ($Demo) { Start-SceneDemo } else { Show-Helper 'hello' }
    # A catalog older than half a day misses new and moved files: refresh it on start.
    $stale = !(Test-Path -LiteralPath $script:IndexPath) -or ((Get-Item -LiteralPath $script:IndexPath).LastWriteTime -lt (Get-Date).AddHours(-12))
    if ($stale -and $rootsBox.Items.Count -gt 0) { $status.Text = 'Обновляю каталог, чтобы найти и новые файлы…'; $form.Refresh(); Update-Catalog; [void]$searchBox.Focus() }
})
$resetButton.Add_Click({ $searchBox.Clear(); Show-Rows $script:AllFiles; $status.Text = "В каталоге файлов: $($script:AllFiles.Count)" })
$openButton.Add_Click({ Open-SelectedFile })
$folderButton.Add_Click({ Show-SelectedFileInFolder })
$planButton.Add_Click({
    $selectedPaths = @(Get-SelectedPaths)
    if ($selectedPaths.Count -gt 0) { Write-OrganizationPlan $selectedPaths; return }
    $desktopRoot = [Environment]::GetFolderPath([Environment+SpecialFolder]::Desktop)
    if (!$desktopRoot) { [void][LibrarianApp.LibBox]::Show('Не удалось определить путь к Рабочему столу.'); return }
    $desktopPrefix = $desktopRoot.TrimEnd('\') + '\'
    $paths = @($script:AllFiles | Where-Object { ([string]$_.Path).StartsWith($desktopPrefix, [StringComparison]::OrdinalIgnoreCase) } | ForEach-Object { [string]$_.Path })
    if ($paths.Count -eq 0) { [void][LibrarianApp.LibBox]::Show('В каталоге нет файлов с Рабочего стола. Нажми «Обновить каталог» и проверь, что Рабочий стол есть в списке папок.'); return }
    Write-OrganizationPlan $paths
})
$cleanupButton.Add_Click({
    if (!$script:HashScanComplete) { [void][LibrarianApp.LibBox]::Show('Сначала нажми «Найти дубликаты».'); return }
    $plan = New-Object System.Collections.Generic.List[object]
    $planLines = New-Object System.Collections.Generic.List[string]
    $planLines.Add('ПЛАН УБОРКИ ТОЧНЫХ ДУБЛИКАТОВ')
    $planLines.Add('Сравнение выполняется по SHA-256. Похожие имена и разные версии файлов не затрагиваются.')
    $planLines.Add('Лишние экземпляры будут отправлены в Корзину Windows; безвозвратного удаления нет.')
    $planLines.Add('')
    $totalBytes = [long]0
    foreach ($group in $script:DuplicateGroups) {
        $verified = New-Object System.Collections.Generic.List[object]
        foreach ($entry in $group.Entries) {
            try {
                if ((Test-Path -LiteralPath $entry.Path -PathType Leaf) -and (Get-FileHash -LiteralPath $entry.Path -Algorithm SHA256 -ErrorAction Stop).Hash -eq $group.Hash) {
                    $verified.Add($entry)
                }
            } catch { }
        }
        if ($verified.Count -lt 2) { continue }
        $ordered = @(Get-KeeperFirst $verified)
        $keeper = $ordered[0]
        $planLines.Add(('Оставить: {0}' -f $keeper.Path))
        foreach ($candidate in @($ordered | Select-Object -Skip 1)) {
            $plan.Add([pscustomobject]@{ Path=$candidate.Path; Hash=$group.Hash; SizeBytes=[long]$group.SizeBytes; Keeper=$keeper.Path })
            $planLines.Add(('В Корзину: {0}' -f $candidate.Path))
            $totalBytes += [long]$group.SizeBytes
        }
        $planLines.Add('')
    }
    if ($plan.Count -eq 0) { [void][LibrarianApp.LibBox]::Show('Подтверждённых точных дублей не осталось. Обнови каталог и повтори проверку.'); return }
    $planLines | Set-Content -LiteralPath $script:CleanupLogPath -Encoding UTF8
    $summary = "Найдено точных копий для Корзины: $($plan.Count)`nМожно освободить примерно: $(Format-Bytes $totalBytes)`n`nВ каждой группе останется один файл. Если он есть вне «Загрузок», будет оставлена эта копия; иначе останется самая короткая по пути папки.`n`nБудут затронуты только папки из списка программы. План: $script:CleanupLogPath`n`nОтправить лишние экземпляры в Корзину?"
    if ([LibrarianApp.LibBox]::Show($summary, 'Подтверждение уборки', 'YesNo', 'Warning') -ne 'Yes') { return }
    $sent = 0; $skipped = 0; $sentBytes = [long]0
    $resultLines = New-Object System.Collections.Generic.List[string]
    $resultLines.AddRange([string[]]$planLines.ToArray())
    $resultLines.Add('ИТОГ')
    $keeperState = @{}
    foreach ($candidate in $plan) {
        try {
            if (!$keeperState.ContainsKey($candidate.Keeper)) {
                $keeperState[$candidate.Keeper] = $false
                if (Test-Path -LiteralPath $candidate.Keeper -PathType Leaf) {
                    $keeperState[$candidate.Keeper] = ((Get-FileHash -LiteralPath $candidate.Keeper -Algorithm SHA256 -ErrorAction Stop).Hash -eq $candidate.Hash)
                }
            }
            if (!$keeperState[$candidate.Keeper]) { $skipped++; $resultLines.Add(('ПРОПУЩЕН (сохраняемый экземпляр изменился или недоступен): {0}' -f $candidate.Path)); continue }
            if (!(Test-Path -LiteralPath $candidate.Path -PathType Leaf)) { $skipped++; $resultLines.Add(('ПРОПУЩЕН (файл не найден): {0}' -f $candidate.Path)); continue }
            if ((Get-FileHash -LiteralPath $candidate.Path -Algorithm SHA256 -ErrorAction Stop).Hash -ne $candidate.Hash) { $skipped++; $resultLines.Add(('ПРОПУЩЕН (содержимое изменилось): {0}' -f $candidate.Path)); continue }
            # UIOption has OnlyErrorDialogs and AllDialogs; DoNothing is not a valid enum value.
            [Microsoft.VisualBasic.FileIO.FileSystem]::DeleteFile($candidate.Path, [Microsoft.VisualBasic.FileIO.UIOption]::OnlyErrorDialogs, [Microsoft.VisualBasic.FileIO.RecycleOption]::SendToRecycleBin, [Microsoft.VisualBasic.FileIO.UICancelOption]::ThrowException)
            $sent++; $sentBytes += [long]$candidate.SizeBytes; $resultLines.Add(('В КОРЗИНЕ: {0}' -f $candidate.Path))
        } catch { $skipped++; $resultLines.Add(('ОШИБКА: {0} | {1}' -f $candidate.Path,$_.Exception.Message)) }
    }
    $resultLines.Add(('Отправлено в Корзину: {0}; пропущено/ошибок: {1}; освобождено приблизительно: {2}' -f $sent,$skipped,(Format-Bytes $sentBytes)))
    $resultLines | Set-Content -LiteralPath $script:CleanupLogPath -Encoding UTF8
    [void][LibrarianApp.LibBox]::Show("Отправлено в Корзину: $sent`nПропущено или ошибок: $skipped`nЖурнал: $script:CleanupLogPath")
    Reset-DuplicateScan
    $status.Text = 'Уборка завершена. Перед новой проверкой нажми «Обновить каталог».'
})
$organizeButton.Add_Click({ Invoke-OrganizationPlan })
$excelPlanButton.Add_Click({ Open-PlanInExcel })
$undoMoveButton.Add_Click({ Undo-LastMove })
$dupeButton.Add_Click({
    if ($script:AllFiles.Count -eq 0) { [void][LibrarianApp.LibBox]::Show('Сначала обнови каталог.'); return }
    Reset-DuplicateScan
    $dupeButton.Enabled = $false; $status.Text = 'Сравниваю содержимое файлов; это может занять время…'; $form.Refresh()
    $byHash = @{}
    $sizeGroups = @($script:AllFiles | Group-Object { [long]$_.SizeBytes })
    foreach ($sizeGroup in $sizeGroups) {
        if ($sizeGroup.Count -lt 2) { continue }
        foreach ($entry in $sizeGroup.Group) {
            if (!(Test-Path -LiteralPath $entry.Path -PathType Leaf)) { continue }
            try {
                if (!$entry.PSObject.Properties['SizeBytes']) {
                    $length = (Get-Item -LiteralPath $entry.Path -ErrorAction Stop).Length
                    $entry | Add-Member -NotePropertyName SizeBytes -NotePropertyValue ([long]$length)
                }
                $hash = (Get-FileHash -LiteralPath $entry.Path -Algorithm SHA256 -ErrorAction Stop).Hash
                $entry.Hash = $hash
                if (!$byHash.ContainsKey($hash)) { $byHash[$hash] = New-Object System.Collections.Generic.List[object] }
                $byHash[$hash].Add($entry)
            } catch { }
        }
    }
    $script:DuplicateGroups = @(
        foreach ($pair in $byHash.GetEnumerator()) {
            if ($pair.Value.Count -gt 1) {
                $entries = @($pair.Value | Sort-Object Path)
                [pscustomobject]@{ Hash=$pair.Key; Entries=$entries; Count=$entries.Count; SizeBytes=[long]$entries[0].SizeBytes }
            }
        }
    )
    $script:HashScanComplete = $true
    $cleanupButton.Enabled = $true
    $dupeRows = @($script:DuplicateGroups | ForEach-Object { $_.Entries })
    $extraCopies = 0; $possibleBytes = [long]0
    foreach ($group in $script:DuplicateGroups) {
        $extraCopies += ($group.Count - 1)
        $possibleBytes += [long]($group.SizeBytes * ($group.Count - 1))
    }
    Show-Rows $dupeRows
    $status.Text = "Групп: $($script:DuplicateGroups.Count); записей: $($dupeRows.Count); лишних копий: $extraCopies; до освобождения: $(Format-Bytes $possibleBytes)."
    $dupeButton.Enabled = $true
    if ($dupeRows.Count -eq 0) { [void][LibrarianApp.LibBox]::Show('Одинаковых файлов в выбранных папках не найдено.') }
})
$reportButton.Add_Click({
    if (!$script:HashScanComplete) { [void][LibrarianApp.LibBox]::Show('Сначала нажми «Найти дубликаты».'); return }
    $totalRows = 0; $extraCopies = 0; $possibleBytes = [long]0
    $lines = New-Object System.Collections.Generic.List[string]
    $lines.Add('ОТЧЁТ О ТОЧНЫХ ДУБЛИКАТАХ')
    $lines.Add('Сканировались только папки, перечисленные в окне программы.')
    $lines.Add('Файлы не удалялись и не перемещались.')
    $lines.Add('')
    foreach ($group in ($script:DuplicateGroups | Sort-Object SizeBytes -Descending)) {
        $totalRows += $group.Count
        $extra = $group.Count - 1
        $extraCopies += $extra
        $possibleBytes += [long]($group.SizeBytes * $extra)
        $sample = $group.Entries[0]
        $lines.Add(('Группа: {0} | файлов: {1} | размер одного: {2} | можно освободить максимум: {3}' -f $sample.Name,$group.Count,(Format-Bytes $group.SizeBytes),(Format-Bytes ([long]($group.SizeBytes * $extra)))))
        foreach ($entry in $group.Entries) {
            $tag = 'кандидат в группу точных дублей; решение будет по SHA-256'
            $lines.Add(('  [{0}] {1}' -f $tag,$entry.Path))
        }
        $lines.Add('')
    }
    $summary = @(
        "Групп одинаковых файлов: $($script:DuplicateGroups.Count)",
        "Файлов в этих группах: $totalRows",
        "Лишних копий при сохранении одной в каждой группе: $extraCopies",
        "Потенциально освободится при удалении всех лишних копий: $(Format-Bytes $possibleBytes)",
        "Уборщик будет оставлять один проверенный экземпляр и отправлять остальные точные копии в Корзину после подтверждения."
    )
    $lines.Insert(4,($summary -join "`r`n"))
    $lines | Set-Content -LiteralPath $script:DuplicateReportPath -Encoding UTF8
    Start-Process notepad.exe -ArgumentList "`"$script:DuplicateReportPath`""
    $status.Text = "Отчёт сохранён. Лишних копий: $extraCopies; потенциально: $(Format-Bytes $possibleBytes)."
})
# Self-update: when this file is changed (improvements), restart into the new version as soon as no dialog is open.
$script:ScriptPath = $PSCommandPath
$script:ScriptStamp = (Get-Item -LiteralPath $PSCommandPath).LastWriteTime
$script:ReloadTimer = New-Object System.Windows.Forms.Timer
$script:ReloadTimer.Interval = 5000
$script:ReloadTimer.Add_Tick({
    try { $now = (Get-Item -LiteralPath $script:ScriptPath).LastWriteTime } catch { return }
    if ($now -le $script:ScriptStamp) { return }
    # Wait while any dialog is open or a job is running; and give the editor a moment to finish writing.
    if ([System.Windows.Forms.Application]::OpenForms.Count -gt 1 -or !$promptsButton.Enabled -or !$weeklyButton.Enabled) { return }
    if (((Get-Date) - $now).TotalSeconds -lt 4) { return }
    $errs = $null; $tok = $null
    [void][System.Management.Automation.Language.Parser]::ParseFile($script:ScriptPath, [ref]$tok, [ref]$errs)
    if ($errs.Count -gt 0) { $script:ScriptStamp = $now; return }   # a broken edit is never started
    $script:ReloadTimer.Stop()
    $answerLabel.Text = 'Меня улучшили — перезапускаюсь…'; $form.Refresh()
    Start-Process -FilePath "$env:SystemRoot\System32\WindowsPowerShell\v1.0\powershell.exe" -ArgumentList ('-NoProfile -STA -ExecutionPolicy Bypass -WindowStyle Hidden -File "' + $script:ScriptPath + '"') -WorkingDirectory $PSScriptRoot
    $script:Closing = $true
    $form.Close()
})
$script:ReloadTimer.Start()
# Goodbye: he waves «Заходите ещё!» and the window closes after the scene (only with Arthur's voice on).
$script:Closing = $false
$form.Add_FormClosing({
    if ($script:Closing -or !$voiceCheck.Checked -or !(Test-Path -LiteralPath (Join-Path $script:HelperDir 'bye.gif'))) { return }
    $_.Cancel = $true; $script:Closing = $true
    Show-Helper 'bye'
    $closeTimer = New-Object System.Windows.Forms.Timer; $closeTimer.Interval = 3600
    $closeTimer.Add_Tick({ $this.Stop(); $form.Close() })
    $closeTimer.Start()
})
# ---- Demo: every scene of the Librarian in turn, with his voice and a note when it happens ----
$script:DemoScenes = @(
    @('hello',  'Сцена 1 из 8 — «Доброе утро! Сколько же у нас сегодня книг…»', 'Когда вы открываете Библиотекаря.'),
    @('search', 'Сцена 2 из 8 — «Так-так… Где же она у нас стояла?»', 'Когда ничего не нашлось, и пока идёт долгая работа (сбор промптов, разбор картинок).'),
    @('found',  'Сцена 3 из 8 — «Ага! Вот ты где!»', 'Когда файл нашёлся, и когда новые промпты разобраны.'),
    @('secret', 'Сцена 4 из 8 — «Тсс… Это наш секрет.»', 'Когда удаляете папку, а внутри похоже на документы или пароли. И в окнах-предупреждениях.'),
    @('sweep',  'Сцена 5 из 8 — «Так, наведём порядок! Ненужное — в корзину, нужное — на полочку.»', 'Еженедельная уборка и удаление в Корзину.'),
    @('dust',   'Сцена 6 из 8 — «Фух! Ха-ха… Какая красота!»', 'Когда уборка закончена, папка создана, картинки разобраны.'),
    @('write',  'Сцена 7 из 8 — «Так, запишем… Новая папка готова, ставлю её на полочку!»', 'Когда создаёте новую папку: пока вы пишете название, он пишет пером, а потом говорит это.'),
    @('bye',    'Сцена 8 из 8 — «Заходите ещё! Я всегда найду нужную книгу.»', 'Когда вы закрываете окно.'))
$script:DemoTimer = New-Object System.Windows.Forms.Timer
$script:DemoTimer.Add_Tick({
    $script:DemoTimer.Stop()
    $script:DemoStep++
    if ($script:DemoStep -ge $script:DemoScenes.Count) {
        $answerLabel.Text = "Это все сцены Библиотекаря.`r`n`r`nМежду сценами он спокойно стоит у полки с книгой. Показать ещё раз: «Ещё ▾» → «Показать все сцены»."
        Show-HelperIdle; return
    }
    $s = $script:DemoScenes[$script:DemoStep]
    $answerLabel.Text = $s[1] + "`r`n`r`n" + $s[2]
    Show-Helper $s[0]
    $script:DemoTimer.Interval = [int]$script:HelperLengthMs[$s[0]] + 900
    $script:DemoTimer.Start()
})
function Start-SceneDemo {
    $script:DemoStep = -1
    $script:DemoTimer.Interval = 300
    $script:DemoTimer.Start()
}
if ($script:FullStage) { . $script:FullOverrides }
[LibrarianApp.LibBox]::Impl = [Func[string,string,string,string,string]]{ param($text, $caption, $buttons, $icon) Show-LibDialog $text $caption $buttons $icon }
# The table may stay hidden all the time; selection in a ListView works only once its window handle exists.
[void]$grid.Handle
Show-HelperIdle
[void]$form.ShowDialog()
