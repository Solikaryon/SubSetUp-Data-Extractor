using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Windows.Forms;
using System.Drawing;
using OfficeOpenXml;
using System.Drawing.Drawing2D;
using System.Runtime.InteropServices;

namespace ExtractorSubSetUp
{
    internal static class NativeFolderPicker
    {
        private const uint SIGDN_FILESYSPATH = 0x80058000;

        [Flags]
        private enum FOS : uint
        {
            PICKFOLDERS = 0x00000020,
            FORCEFILESYSTEM = 0x00000040,
            ALLOWMULTISELECT = 0x00000200,
            PATHMUSTEXIST = 0x00000800,
        }

        [ComImport]
        [Guid("b4db1657-70d7-485e-8e3e-6fcb5a5c1802")]
        [InterfaceType(ComInterfaceType.InterfaceIsIUnknown)]
        private interface IModalWindow
        {
            [PreserveSig]
            int Show(IntPtr parent);
        }

        [ComImport]
        [Guid("42f85136-db7e-439c-85f6-1bea794e4e99")]
        [InterfaceType(ComInterfaceType.InterfaceIsIUnknown)]
        private interface IFileDialog : IModalWindow
        {
            [PreserveSig] int SetFileTypes(uint cFileTypes, IntPtr rgFilterSpec);
            [PreserveSig] int SetFileTypeIndex(uint iFileType);
            [PreserveSig] int GetFileTypeIndex(out uint piFileType);
            [PreserveSig] int Advise(IntPtr pfde, out uint pdwCookie);
            [PreserveSig] int Unadvise(uint dwCookie);
            [PreserveSig] int SetOptions(FOS fos);
            [PreserveSig] int GetOptions(out FOS pfos);
            [PreserveSig] int SetDefaultFolder(IShellItem psi);
            [PreserveSig] int SetFolder(IShellItem psi);
            [PreserveSig] int GetFolder(out IShellItem ppsi);
            [PreserveSig] int GetCurrentSelection(out IShellItem ppsi);
            [PreserveSig] int SetFileName([MarshalAs(UnmanagedType.LPWStr)] string pszName);
            [PreserveSig] int GetFileName([MarshalAs(UnmanagedType.LPWStr)] out string pszName);
            [PreserveSig] int SetTitle([MarshalAs(UnmanagedType.LPWStr)] string pszTitle);
            [PreserveSig] int SetOkButtonLabel([MarshalAs(UnmanagedType.LPWStr)] string pszText);
            [PreserveSig] int SetFileNameLabel([MarshalAs(UnmanagedType.LPWStr)] string pszLabel);
            [PreserveSig] int GetResult(out IShellItem ppsi);
            [PreserveSig] int AddPlace(IShellItem psi, uint fdap);
            [PreserveSig] int SetDefaultExtension([MarshalAs(UnmanagedType.LPWStr)] string pszDefaultExt);
            [PreserveSig] int Close(int hr);
            [PreserveSig] int SetClientGuid(ref Guid guid);
            [PreserveSig] int ClearClientData();
            [PreserveSig] int SetFilter(IntPtr pFilter);
        }

        [ComImport]
        [Guid("d57c7288-d4ad-4768-be02-9d969532d960")]
        [InterfaceType(ComInterfaceType.InterfaceIsIUnknown)]
        private interface IFileOpenDialog : IFileDialog
        {
            [PreserveSig] int GetResults(out IShellItemArray ppenum);
            [PreserveSig] int GetSelectedItems(out IShellItemArray ppsai);
        }

        [ComImport]
        [Guid("DC1C5A9C-E88A-4DDE-A5A1-60F82A20AEF7")]
        private class FileOpenDialogCom
        {
        }

        [ComImport]
        [Guid("B63EA76D-1F85-456F-A19C-48159EFA858B")]
        [InterfaceType(ComInterfaceType.InterfaceIsIUnknown)]
        private interface IShellItemArray
        {
            [PreserveSig] int BindToHandler(IntPtr pbc, ref Guid bhid, ref Guid riid, out IntPtr ppvOut);
            [PreserveSig] int GetPropertyStore(int flags, ref Guid riid, out IntPtr ppv);
            [PreserveSig] int GetPropertyDescriptionList(ref Guid keyType, ref Guid riid, out IntPtr ppv);
            [PreserveSig] int GetAttributes(uint AttribFlags, uint sfgaoMask, out uint psfgaoAttribs);
            [PreserveSig] int GetCount(out uint pdwNumItems);
            [PreserveSig] int GetItemAt(uint dwIndex, out IShellItem ppsi);
            [PreserveSig] int EnumItems(out IntPtr ppenumShellItems);
        }

        [ComImport]
        [Guid("43826D1E-E718-42EE-BC55-A1E261C37BFE")]
        [InterfaceType(ComInterfaceType.InterfaceIsIUnknown)]
        private interface IShellItem
        {
            [PreserveSig] int BindToHandler(IntPtr pbc, ref Guid bhid, ref Guid riid, out IntPtr ppv);
            [PreserveSig] int GetParent(out IShellItem ppsi);
            [PreserveSig] int GetDisplayName(uint sigdnName, out IntPtr ppszName);
            [PreserveSig] int GetAttributes(uint sfgaoMask, out uint psfgaoAttribs);
            [PreserveSig] int Compare(IShellItem psi, uint hint, out int piOrder);
        }

        public static string[] PickFolders(IntPtr ownerHwnd)
        {
            IFileOpenDialog? dialog = null;
            IShellItemArray? results = null;

            try
            {
                dialog = (IFileOpenDialog)new FileOpenDialogCom();

                dialog.GetOptions(out var current);
                var options = current | FOS.PICKFOLDERS | FOS.FORCEFILESYSTEM | FOS.ALLOWMULTISELECT | FOS.PATHMUSTEXIST;
                dialog.SetOptions(options);
                dialog.SetTitle("Select folders");

                var hr = dialog.Show(ownerHwnd);
                if (hr != 0)
                    return Array.Empty<string>();

                dialog.GetResults(out results);
                results.GetCount(out var count);

                var folders = new List<string>((int)count);
                for (uint i = 0; i < count; i++)
                {
                    results.GetItemAt(i, out var item);
                    try
                    {
                        item.GetDisplayName(SIGDN_FILESYSPATH, out var ppsz);
                        try
                        {
                            var path = Marshal.PtrToStringUni(ppsz);
                            if (!string.IsNullOrWhiteSpace(path))
                                folders.Add(path);
                        }
                        finally
                        {
                            if (ppsz != IntPtr.Zero)
                                Marshal.FreeCoTaskMem(ppsz);
                        }
                    }
                    finally
                    {
                        if (item != null)
                            Marshal.ReleaseComObject(item);
                    }
                }

                return folders.Distinct(StringComparer.OrdinalIgnoreCase).ToArray();
            }
            finally
            {
                if (results != null)
                    Marshal.ReleaseComObject(results);
                if (dialog != null)
                    Marshal.ReleaseComObject(dialog);
            }
        }
    }

    public static class FolderBrowserMultiSelect
    {
        public static string[] SelectFolders()
        {
            using (var form = new MultiFolderDialog())
            {
                if (form.ShowDialog() == DialogResult.OK)
                    return form.SelectedFolders.ToArray();
            }

            return Array.Empty<string>();
        }

        public static string[] SelectFoldersNative(IWin32Window owner)
        {
            // Abrir directamente el selector alternativo (sin avisos) para evitar mensajes/errores
            // y permitir agregar múltiples carpetas con portapapeles o arrastrar y soltar.
            return SelectFolders();
        }

        public static IEnumerable<string> ResolveFeederSetupFiles(string baseDir)
        {
            var results = new HashSet<string>(StringComparer.OrdinalIgnoreCase);

            foreach (var root in GetSearchRoots(baseDir))
            {
                if (!Directory.Exists(root))
                    continue;

                var directPath = Path.Combine(root, "FeederSetup.csv");
                if (File.Exists(directPath))
                {
                    results.Add(directPath);
                    continue;
                }

                var dataDir = FindChildDirectory(root, "DATA");
                if (dataDir != null)
                {
                    var dataPath = Path.Combine(dataDir, "FeederSetup.csv");
                    if (File.Exists(dataPath))
                    {
                        results.Add(dataPath);
                        continue;
                    }
                }

                foreach (var match in Directory.GetFiles(root, "FeederSetup*.csv", SearchOption.AllDirectories))
                    results.Add(match);
            }

            return results;
        }

        private static IEnumerable<string> GetSearchRoots(string selectedDir)
        {
            yield return selectedDir;

            var dirName = Path.GetFileName(selectedDir.TrimEnd(Path.DirectorySeparatorChar, Path.AltDirectorySeparatorChar));
            if (string.Equals(dirName, "Data", StringComparison.OrdinalIgnoreCase))
            {
                var parent = Directory.GetParent(selectedDir)?.FullName;
                if (!string.IsNullOrWhiteSpace(parent))
                    yield return parent;
            }
        }

        private static string? FindChildDirectory(string root, string targetName)
        {
            try
            {
                foreach (var dir in Directory.GetDirectories(root, "*", SearchOption.TopDirectoryOnly))
                {
                    var name = Path.GetFileName(dir);
                    if (string.Equals(name, targetName, StringComparison.OrdinalIgnoreCase))
                        return dir;
                }
            }
            catch
            {
                return null;
            }

            return null;
        }
    }

    public class MultiFolderDialog : Form
    {
        private ListBox lstFolders = null!;
        private Button btnAddFolder = null!;
        private Button btnAddClipboard = null!;
        private Button btnRemove = null!;
        private Button btnOK = null!;
        private Button btnCancel = null!;

        public List<string> SelectedFolders { get; private set; } = new();

        public MultiFolderDialog()
        {
            InitializeComponent();
        }

        private void InitializeComponent()
        {
            this.Text = "Select folders";
            this.Width = 520;
            this.Height = 420;
            this.StartPosition = FormStartPosition.CenterScreen;
            this.FormBorderStyle = FormBorderStyle.FixedDialog;
            this.MaximizeBox = false;
            this.MinimizeBox = false;
            this.AllowDrop = true;
            this.BackColor = SystemColors.Control;
            this.ForeColor = SystemColors.ControlText;

            var label = new Label();
            label.Text = "Selected folders:";
            label.Left = 10;
            label.Top = 10;
            label.Width = 300;
            label.Height = 20;
            this.Controls.Add(label);

            lstFolders = new ListBox();
            lstFolders.Left = 10;
            lstFolders.Top = 35;
            lstFolders.Width = 480;
            lstFolders.Height = 300;
            lstFolders.SelectionMode = System.Windows.Forms.SelectionMode.One;
            lstFolders.AllowDrop = true;
            lstFolders.DragEnter += (_, e) =>
            {
                if (e.Data != null && e.Data.GetDataPresent(DataFormats.FileDrop))
                    e.Effect = DragDropEffects.Copy;
            };
            lstFolders.DragDrop += (_, e) =>
            {
                if (e.Data == null)
                    return;

                if (e.Data.GetData(DataFormats.FileDrop) is string[] paths)
                {
                    AddFolders(paths);
                }
            };
            this.Controls.Add(lstFolders);

            btnAddFolder = new Button();
            btnAddFolder.Text = "+ Add folder";
            btnAddFolder.Left = 10;
            btnAddFolder.Top = 345;
            btnAddFolder.Width = 150;
            btnAddFolder.Height = 30;
            btnAddFolder.UseVisualStyleBackColor = false;
            btnAddFolder.BackColor = Color.AliceBlue;
            btnAddFolder.ForeColor = SystemColors.ControlText;
            btnAddFolder.Click += (_, __) => AddFolder();
            this.Controls.Add(btnAddFolder);

            btnAddClipboard = new Button();
            btnAddClipboard.Text = "+ From clipboard";
            btnAddClipboard.Left = 165;
            btnAddClipboard.Top = 345;
            btnAddClipboard.Width = 140;
            btnAddClipboard.Height = 30;
            btnAddClipboard.UseVisualStyleBackColor = false;
            btnAddClipboard.BackColor = Color.Lavender;
            btnAddClipboard.ForeColor = SystemColors.ControlText;
            btnAddClipboard.Click += (_, __) => AddFromClipboard();
            this.Controls.Add(btnAddClipboard);

            btnRemove = new Button();
            btnRemove.Text = "- Remove";
            btnRemove.Left = 310;
            btnRemove.Top = 345;
            btnRemove.Width = 110;
            btnRemove.Height = 30;
            btnRemove.UseVisualStyleBackColor = false;
            btnRemove.BackColor = Color.Moccasin;
            btnRemove.ForeColor = SystemColors.ControlText;
            btnRemove.Click += (_, __) => RemoveFolder();
            this.Controls.Add(btnRemove);

            btnOK = new Button();
            btnOK.Text = "OK";
            btnOK.Left = 425;
            btnOK.Top = 345;
            btnOK.Width = 65;
            btnOK.Height = 30;
            btnOK.Click += (_, __) =>
            {
                SelectedFolders = lstFolders.Items.Cast<string>().ToList();
                this.DialogResult = DialogResult.OK;
                this.Close();
            };
            btnOK.UseVisualStyleBackColor = false;
            btnOK.BackColor = Color.PaleGreen; // verde más visible pero suave
            btnOK.ForeColor = SystemColors.ControlText;
            this.Controls.Add(btnOK);

            btnCancel = new Button();
            btnCancel.Text = "Cancel";
            btnCancel.Left = 410;
            btnCancel.Top = 380;
            btnCancel.Width = 80;
            btnCancel.Height = 30;
            btnCancel.DialogResult = DialogResult.Cancel;
            btnCancel.UseVisualStyleBackColor = false;
            btnCancel.BackColor = Color.LightCoral; // rojo más visible pero suave
            btnCancel.ForeColor = SystemColors.ControlText;
            this.Controls.Add(btnCancel);

            this.AcceptButton = btnOK;
            this.CancelButton = btnCancel;
        }

        private void AddFolder()
        {
            using (var fbd = new FolderBrowserDialog())
            {
                fbd.Description = "Select a folder";
                if (fbd.ShowDialog() == DialogResult.OK && !string.IsNullOrWhiteSpace(fbd.SelectedPath))
                {
                    AddFolders(new[] { fbd.SelectedPath });
                }
            }
        }

        private void AddFromClipboard()
        {
            try
            {
                if (Clipboard.ContainsFileDropList())
                {
                    var list = Clipboard.GetFileDropList();
                    var paths = list.Cast<string>().ToArray();
                    AddFolders(paths);
                    return;
                }

                if (Clipboard.ContainsText())
                {
                    var text = Clipboard.GetText();
                    var paths = text
                        .Split(new[] { "\r\n", "\n", "\r", "\t", ";" }, StringSplitOptions.RemoveEmptyEntries)
                        .Select(p => p.Trim().Trim('"'))
                        .Where(p => !string.IsNullOrWhiteSpace(p))
                        .ToArray();
                    AddFolders(paths);
                    return;
                }

                MessageBox.Show(
                    "The clipboard does not contain any paths.\n\nTip: In File Explorer, select multiple folders and press Ctrl+C.",
                    "Information",
                    MessageBoxButtons.OK,
                    MessageBoxIcon.Information);
            }
            catch (Exception ex)
            {
                MessageBox.Show(
                    "Could not read the clipboard.\n\nDetails: " + ex.Message,
                    "Warning",
                    MessageBoxButtons.OK,
                    MessageBoxIcon.Warning);
            }
        }

        private void AddFolders(IEnumerable<string> paths)
        {
            foreach (var raw in paths)
            {
                var p = (raw ?? string.Empty).Trim().Trim('"');
                if (string.IsNullOrWhiteSpace(p))
                    continue;

                if (!Directory.Exists(p))
                    continue;

                if (!lstFolders.Items.Contains(p))
                    lstFolders.Items.Add(p);
            }
        }

        private void RemoveFolder()
        {
            if (lstFolders.SelectedIndex >= 0)
                lstFolders.Items.RemoveAt(lstFolders.SelectedIndex);
        }
    }

    public class Form1 : Form
    {
        private Label labelJobFolder = null!;
        private Button btnSeleccionar = null!;
        private Button btnSeleccionarCarpetas = null!;
        private TabControl tabControl = null!;

        public Form1()
        {
            InicializarComponentes();
        }

        private void InicializarComponentes()
        {
            this.Text = "Output Report Data Extractor";
            this.Width = 900;
            this.Height = 500;
            this.StartPosition = FormStartPosition.CenterScreen;
            this.BackColor = SystemColors.Control;
            this.ForeColor = SystemColors.ControlText;

            btnSeleccionar = new Button();
            btnSeleccionar.Text = "Select CSV files";
            btnSeleccionar.Top = 10;
            btnSeleccionar.Left = 10;
            btnSeleccionar.Width = 200;
            btnSeleccionar.UseVisualStyleBackColor = false;
            btnSeleccionar.BackColor = Color.LightSkyBlue;
            btnSeleccionar.ForeColor = SystemColors.ControlText;
            btnSeleccionar.Click += SeleccionarArchivo;

            btnSeleccionarCarpetas = new Button();
            btnSeleccionarCarpetas.Text = "Generate folder with files";
            btnSeleccionarCarpetas.Top = 10;
            btnSeleccionarCarpetas.Left = 220;
            btnSeleccionarCarpetas.Width = 170;
            btnSeleccionarCarpetas.UseVisualStyleBackColor = false;
            btnSeleccionarCarpetas.BackColor = Color.PaleGreen;
            btnSeleccionarCarpetas.ForeColor = SystemColors.ControlText;
            btnSeleccionarCarpetas.Click += SeleccionarCarpeta;

            var btnExportar = new Button();
            btnExportar.Text = "Export to Excel";
            btnExportar.Top = 10;
            btnExportar.Left = 400;
            btnExportar.Width = 150;
            btnExportar.UseVisualStyleBackColor = false;
            btnExportar.BackColor = Color.Khaki;
            btnExportar.ForeColor = SystemColors.ControlText;
            btnExportar.Click += ExportarExcel;

            labelJobFolder = new Label();
            labelJobFolder.Text = "JobFolder: ";
            labelJobFolder.Top = 50;
            labelJobFolder.Left = 10;
            labelJobFolder.Width = 800;
            labelJobFolder.Font = new System.Drawing.Font("Arial", 10, System.Drawing.FontStyle.Bold);

            tabControl = new TabControl();
            tabControl.ShowToolTips = true;
            tabControl.Top = 80;
            tabControl.Left = 10;
            tabControl.Width = 850;
            tabControl.Height = 350;
            EnableTabClosing(tabControl);
            tabControl.SelectedIndexChanged += (s, ev) =>
            {
                if (tabControl.SelectedTab?.Tag is Tuple<string,string> t)
                    labelJobFolder.Text = $"JobFolder: {t.Item1}    JobName: {t.Item2}";
                else
                    labelJobFolder.Text = $": {tabControl.TabPages.Count}";
            };

            this.Controls.Add(btnSeleccionar);
            this.Controls.Add(btnSeleccionarCarpetas);
            this.Controls.Add(btnExportar);
            this.Controls.Add(labelJobFolder);
            this.Controls.Add(tabControl);
        }

        private void SeleccionarArchivo(object? sender, EventArgs e)
        {
            using OpenFileDialog openFileDialog = new OpenFileDialog();
            openFileDialog.Filter = "CSV files (*.csv)|*.csv|All files (*.*)|*.*";
            openFileDialog.Title = "Select files";
            openFileDialog.Multiselect = true;

            if (openFileDialog.ShowDialog() != DialogResult.OK)
                return;

            try
            {
                var files = openFileDialog.FileNames;
                if (files == null || files.Length == 0)
                    return;

                ProcesarArchivos(files);
            }
            catch (Exception ex)
            {
                MessageBox.Show($"An error occurred:\n{ex.Message}", "Error",
                    MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        private void SeleccionarCarpeta(object? sender, EventArgs e)
        {
            try
            {
                // 1) Selecciona múltiples carpetas origen (Explorador de archivos con Ctrl/Shift)
                var selectedDirs = FolderBrowserMultiSelect.SelectFoldersNative(this);
                if (selectedDirs == null || selectedDirs.Length == 0)
                    return;

                // 2) Selecciona una carpeta destino
                string? destFolder = null;
                using (var fbd = new FolderBrowserDialog())
                {
                    fbd.Description = "Select the DESTINATION folder where the FeederSetup*.csv files will be saved";
                    if (fbd.ShowDialog() != DialogResult.OK || string.IsNullOrWhiteSpace(fbd.SelectedPath))
                        return;
                    destFolder = fbd.SelectedPath;
                }

                if (string.IsNullOrWhiteSpace(destFolder) || !Directory.Exists(destFolder))
                {
                    MessageBox.Show("The destination folder is not valid.", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    return;
                }

                var copied = 0;
                var errors = new List<string>();
                var foundAny = false;

                foreach (var selectedDir in selectedDirs)
                {
                    var normalizedDir = NormalizePath(selectedDir);
                    if (string.IsNullOrWhiteSpace(normalizedDir))
                        continue;

                    if (!Directory.Exists(normalizedDir))
                    {
                        errors.Add($"Not found: {normalizedDir}");
                        continue;
                    }

                    var rootName = new DirectoryInfo(normalizedDir).Name;
                    IEnumerable<string> feederFiles;
                    try
                    {
                        feederFiles = FolderBrowserMultiSelect.ResolveFeederSetupFiles(normalizedDir)
                            .Distinct(StringComparer.OrdinalIgnoreCase)
                            .ToArray();
                    }
                    catch (Exception ex)
                    {
                        errors.Add($"{normalizedDir}: {ex.Message}");
                        continue;
                    }

                    var index = 0;
                    foreach (var srcFile in feederFiles)
                    {
                        foundAny = true;
                        index++;
                        try
                        {
                            var srcName = Path.GetFileName(srcFile);
                            var nameNoExt = Path.GetFileNameWithoutExtension(srcName);
                            var ext = Path.GetExtension(srcName);

                            var hasPrimaryName = TryGetPrimaryNameForSavedFile(srcFile, out var primaryName);

                            var baseName = hasPrimaryName
                                ? primaryName
                                : (index == 1
                                    ? $"{rootName}__{nameNoExt}"
                                    : $"{rootName}__{nameNoExt}__{index}");

                            var safeName = MakeSafeFileName(baseName) + ext;
                            var destPath = GetUniqueFilePath(destFolder, safeName);
                            File.Copy(srcFile, destPath, overwrite: false);
                            copied++;
                        }
                        catch (Exception ex)
                        {
                            errors.Add($"{srcFile}: {ex.Message}");
                        }
                    }
                }

                if (!foundAny)
                {
                    var detail = errors.Count > 0
                        ? "\n\nDetails:\n" + string.Join("\n", errors.Take(10))
                        : string.Empty;
                    MessageBox.Show("No FeederSetup*.csv files were found in the selected folders." + detail,
                        "Information", MessageBoxButtons.OK, MessageBoxIcon.Information);
                    return;
                }

                var msg = $"Done. Copied {copied} file(s) to:\n{destFolder}";
                if (errors.Count > 0)
                    msg += "\n\nSome errors:\n" + string.Join("\n", errors.Take(8));

                MessageBox.Show(msg, "Result", MessageBoxButtons.OK, MessageBoxIcon.Information);
            }
            catch (Exception ex)
            {
                MessageBox.Show($"An error occurred:\n{ex.Message}", "Error",
                    MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        private static string MakeSafeFileName(string name)
        {
            var invalid = Path.GetInvalidFileNameChars();
            var chars = name.Select(c => invalid.Contains(c) ? '_' : c).ToArray();
            var result = new string(chars).Trim();
            return string.IsNullOrWhiteSpace(result) ? "FeederSetup" : result;
        }

        private static bool TryGetPrimaryNameForSavedFile(string srcFile, out string primaryName)
        {
            primaryName = string.Empty;

            try
            {
                // Intenta extraer un identificador estable desde el contenido del CSV.
                // Caso típico: el CSV está delimitado por ',' o '\t' y el JobName contiene metadatos separados por ';'.
                // Ejemplo: "ARSPCA-01488-50_TOP;B0;Top;WOLVERINE;[Top] L715B;Original setup" => "ARSPCA-01488-50_TOP".

                int headerIdx = -1;
                char delimiter = ',';
                int currentIndex = -1;

                foreach (var line in File.ReadLines(srcFile))
                {
                    currentIndex++;
                    if (currentIndex > 200)
                        break;

                    if (string.IsNullOrWhiteSpace(line))
                        continue;

                    if (TryParseHeaders(line, out _, out var sep))
                    {
                        headerIdx = currentIndex;
                        delimiter = sep;
                        continue;
                    }

                    if (headerIdx >= 0 && currentIndex <= headerIdx)
                        continue;

                    // Primera línea de datos después de headers (si existen).
                    var cols = ParseDelimitedLine(line, delimiter)
                        .Select(NormalizeCell)
                        .ToArray();

                    if (cols.Length == 0)
                        continue;

                    string? candidate = null;
                    var c0 = cols.Length > 0 ? cols[0] : null;
                    var c1 = cols.Length > 1 ? cols[1] : null;

                    // Preferir el que parezca ID (contenga '-' o '_')
                    foreach (var c in new[] { c1, c0 })
                    {
                        if (string.IsNullOrWhiteSpace(c))
                            continue;
                        var first = c.Split(';')[0].Trim();
                        if (string.IsNullOrWhiteSpace(first))
                            continue;
                        if (first.Contains('-') || first.Contains('_'))
                        {
                            candidate = first;
                            break;
                        }
                    }

                    // Fallback: usar el primer token del mejor candidato disponible
                    if (string.IsNullOrWhiteSpace(candidate))
                    {
                        var raw = !string.IsNullOrWhiteSpace(c1) ? c1 : c0;
                        if (!string.IsNullOrWhiteSpace(raw))
                            candidate = raw.Split(';')[0].Trim();
                    }

                    if (string.IsNullOrWhiteSpace(candidate))
                        continue;

                    primaryName = candidate;
                    return true;
                }
            }
            catch
            {
                // Ignorar y hacer fallback al nombre por carpeta/archivo.
            }

            return false;
        }

        private static string GetUniqueFilePath(string destFolder, string fileName)
        {
            var baseName = Path.GetFileNameWithoutExtension(fileName);
            var ext = Path.GetExtension(fileName);
            var candidate = Path.Combine(destFolder, fileName);

            var n = 2;
            while (File.Exists(candidate))
            {
                candidate = Path.Combine(destFolder, $"{baseName}_{n}{ext}");
                n++;
            }

            return candidate;
        }

        private void ProcesarArchivos(IEnumerable<string> files)
        {
            // Create loading form
            bool cancelRequested = false;
            using var loadingForm = new Form
            {
                Text = "Loading...",
                Width = 300,
                Height = 120,
                StartPosition = FormStartPosition.CenterScreen,
                ControlBox = true,
                FormBorderStyle = FormBorderStyle.FixedDialog,
                BackColor = SystemColors.Control,
                ForeColor = SystemColors.ControlText
            };
            loadingForm.FormClosing += (_, __) => cancelRequested = true;

            var lblLoading = new Label
            {
                Text = "Please wait, loading data...",
                Dock = DockStyle.Fill,
                TextAlign = ContentAlignment.MiddleCenter,
                Font = new Font("Arial", 10, FontStyle.Bold),
                BackColor = Color.Transparent,
                ForeColor = SystemColors.ControlText
            };
            loadingForm.Controls.Add(lblLoading);
            loadingForm.Show();
            Application.DoEvents();

            // Do not clear existing tabs: always add a new tab per selected file

            foreach (var file in files)
            {
                if (cancelRequested)
                    break;

                try
                {
                    var lineas = File.ReadAllLines(file);

                    string JobFolder = "No encontrado";
                    string JobName = "No encontrado";

                    // buscar encabezados y detectar separador
                    int encabezadosIdx = -1;
                    char delimiter = ',';
                    string[] encabezados = Array.Empty<string>();
                    int idxModule = -1;
                    int idxPart = -1;
                    int idxLocation = -1;
                    int idxSide = -1;

                    for (int i = 0; i < lineas.Length; i++)
                    {
                        if (TryParseHeaders(lineas[i], out var cols, out var sep))
                        {
                            encabezadosIdx = i;
                            delimiter = sep;
                            encabezados = cols;
                            idxModule = Array.FindIndex(encabezados, x => string.Equals(x, "ModuleNumber", StringComparison.OrdinalIgnoreCase));
                            idxPart = Array.FindIndex(encabezados, x => string.Equals(x, "PartNumber", StringComparison.OrdinalIgnoreCase));
                            idxLocation = Array.FindIndex(encabezados, x => string.Equals(x, "Location", StringComparison.OrdinalIgnoreCase));
                            idxSide = Array.FindIndex(encabezados, x => string.Equals(x, "SideNo", StringComparison.OrdinalIgnoreCase));
                            break;
                        }
                    }

                    if (encabezadosIdx == -1)
                    {
                        // añadir pestaña vacía indicando error
                        var tpErr = new TabPage(Path.GetFileName(file));
                        var lbl = new Label() { Text = "Required columns were not found.", Dock = DockStyle.Fill, TextAlign = ContentAlignment.MiddleCenter };
                        tpErr.Controls.Add(lbl);
                        tabControl.TabPages.Add(tpErr);
                        continue;
                    }

                    if (lineas.Length > 1)
                    {
                        var primerDato = ParseDelimitedLine(lineas[1], delimiter).ToArray();
                        // keep backslashes in JobFolder (don't trim '\')
                        JobFolder = primerDato.Length > 0 ? NormalizeCell(primerDato[0]) : JobFolder;
                        JobName = primerDato.Length > 1 ? NormalizeCell(primerDato[1]) : JobName;
                    }

                    var grid = new DataGridView();
                    grid.Dock = DockStyle.Fill;
                    grid.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.Fill;
                    grid.AllowUserToAddRows = false;
                    grid.Columns.Add("JobName", "JobName");
                    grid.Columns.Add("SideNo", "SideNo");
                    grid.Columns.Add("ModuleNumber", "ModuleNumber");
                    grid.Columns.Add("PartNumber", "PartNumber");
                    grid.Columns.Add("Location", "Location");
                    if (grid.Columns["SideNo"] is DataGridViewColumn sideNoColumn)
                        sideNoColumn.Visible = false;

                    for (int i = encabezadosIdx + 1; i < lineas.Length; i++)
                    {
                        if (cancelRequested)
                            break;

                        if (string.IsNullOrWhiteSpace(lineas[i]))
                            continue;

                        var cols = ParseDelimitedLine(lineas[i], delimiter)
                            .Select(NormalizeCell)
                            .ToArray();

                        if (cols.Length <= Math.Max(Math.Max(idxModule, idxPart), Math.Max(idxLocation, idxSide)))
                            continue;

                        var sideNo = (idxSide >= 0 && idxSide < cols.Length) ? cols[idxSide] : string.Empty;

                        grid.Rows.Add(
                            JobName,
                            sideNo,
                            cols[idxModule],
                            cols[idxPart],
                            cols[idxLocation]
                        );
                    }

                    // create top panel with JobFolder / JobName
                    var topPanel = new Panel() { Dock = DockStyle.Top, Height = 26 };
                    var lblJobFolderTab = new Label() { Text = $"JobFolder: {JobFolder}", AutoSize = false, TextAlign = ContentAlignment.MiddleLeft, Dock = DockStyle.Left, Width = 420 };
                    var lblJobNameTab = new Label() { Text = $"JobName: {JobName}", AutoSize = false, TextAlign = ContentAlignment.MiddleRight, Dock = DockStyle.Right, Width = 420 };
                    topPanel.Controls.Add(lblJobFolderTab);
                    topPanel.Controls.Add(lblJobNameTab);

                    var tp = new TabPage();
                    // tab title: filename (parent folder) to distinguish identical names in different folders
                    string parent = Path.GetFileName(Path.GetDirectoryName(file) ?? string.Empty) ?? string.Empty;
                    tp.Text = Path.GetFileName(file) + (string.IsNullOrEmpty(parent) ? "" : $" ({parent})");
                    tp.ToolTipText = file;
                    tp.Tag = Tuple.Create(JobFolder, JobName);
                    tp.Controls.Add(grid);
                    tp.Controls.Add(topPanel);
                    tabControl.TabPages.Add(tp);
                    tabControl.SelectedTab = tp;
                }
                catch (Exception ex)
                {
                    var tpErr = new TabPage(Path.GetFileName(file));
                    var lbl = new Label() { Text = "Error: " + ex.Message, Dock = DockStyle.Fill, TextAlign = ContentAlignment.MiddleCenter };
                    tpErr.Controls.Add(lbl);
                    tabControl.TabPages.Add(tpErr);
                }
            }

            if (!loadingForm.IsDisposed)
                loadingForm.Close();
        }

        private void ExportarExcel(object? sender, EventArgs e)
        {
            if (tabControl.TabPages.Count == 0)
            {
                MessageBox.Show("There is no data to export.", "Information", MessageBoxButtons.OK, MessageBoxIcon.Information);
                return;
            }

            using SaveFileDialog sfd = new SaveFileDialog();
            sfd.Filter = "Excel files (*.xlsx)|*.xlsx";
            sfd.DefaultExt = "xlsx";
            sfd.FileName = $"FeederSetup_{DateTime.Now:yyyyMMdd_HHmmss}.xlsx";

            if (sfd.ShowDialog() != DialogResult.OK)
                return;

            // Create loading form
            bool cancelRequested = false;
            using var loadingForm = new Form
            {
                Text = "Exporting...",
                Width = 320,
                Height = 140,
                StartPosition = FormStartPosition.CenterScreen,
                ControlBox = true,
                FormBorderStyle = FormBorderStyle.FixedDialog,
                BackColor = SystemColors.Control,
                ForeColor = SystemColors.ControlText
            };
            loadingForm.FormClosing += (_, __) => cancelRequested = true;

            var lblLoading = new Label
            {
                Text = "Please wait, exporting data...",
                Dock = DockStyle.Fill,
                TextAlign = ContentAlignment.MiddleCenter,
                Font = new Font("Arial", 10, FontStyle.Bold),
                BackColor = Color.Transparent,
                ForeColor = SystemColors.ControlText
            };
            loadingForm.Controls.Add(lblLoading);
            loadingForm.Show();
            Application.DoEvents();

            try
            {
                OfficeOpenXml.ExcelPackage.LicenseContext = OfficeOpenXml.LicenseContext.NonCommercial;

                using (var package = new OfficeOpenXml.ExcelPackage())
                {
                    // Consolidate data from all tabs
                    // Nota: usar una clave única por pestaña evita mezclar datos cuando hay JobName repetidos.
                    var dataByJobName = new Dictionary<string, List<(string ModuleSlot, List<string> PartNumbers)>>();
                    var allModuleSlots = new HashSet<string>();
                    var usedColumnNames = new HashSet<string>(StringComparer.OrdinalIgnoreCase);

                    foreach (TabPage tab in tabControl.TabPages)
                    {
                        if (cancelRequested)
                            break;

                        // Get JobName from tag
                        var jobName = "Unknown";
                        if (tab.Tag is Tuple<string, string> t)
                            jobName = t.Item2;

                        var columnName = GetUniqueName(jobName, usedColumnNames);

                        // Get the DataGridView from the tab
                        DataGridView? grid = null;
                        foreach (Control ctrl in tab.Controls)
                        {
                            if (ctrl is DataGridView dgv)
                            {
                                grid = dgv;
                                break;
                            }
                        }

                        if (grid == null)
                            continue;

                        if (!dataByJobName.ContainsKey(columnName))
                            dataByJobName[columnName] = new List<(string, List<string>)>();

                        // Group data by ModuleSlot first
                        var tempData = new Dictionary<string, List<string>>();

                        // Extract data from grid
                        for (int row = 0; row < grid.Rows.Count; row++)
                        {
                            if (cancelRequested)
                                break;

                            var moduleNumber = (grid.Rows[row].Cells["ModuleNumber"].Value?.ToString() ?? "").Trim();
                            var location = (grid.Rows[row].Cells["Location"].Value?.ToString() ?? "").Trim();
                            var partNumber = (grid.Rows[row].Cells["PartNumber"].Value?.ToString() ?? "").Trim();

                            // Ya no usar SideNo en la clave
                            var moduleSlot = BuildModuleSlotWithoutSide(moduleNumber, location);
                            
                            if (!tempData.ContainsKey(moduleSlot))
                                tempData[moduleSlot] = new List<string>();
                            
                            if (!string.IsNullOrWhiteSpace(partNumber))
                                tempData[moduleSlot].Add(partNumber);
                            
                            allModuleSlots.Add(moduleSlot);
                        }

                        // Convert to final format
                        foreach (var kvp in tempData)
                        {
                            dataByJobName[columnName].Add((kvp.Key, kvp.Value.Distinct().ToList()));
                        }
                    }

                    // Create consolidated sheet
                    var sheet = package.Workbook.Worksheets.Add("Consolidado");
                    var sheetAlternate = package.Workbook.Worksheets.Add("Alternos");

                    var jobNames = dataByJobName.Keys.ToList();
                    var sortedModuleSlots = SortModuleSlots(allModuleSlots.ToList());

                    // Write headers - Row 1: JobNames, Row 2: Column labels
                    sheet.Cells[2, 1].Value = "Module-Slot";
                    sheet.Cells[2, 1].Style.Font.Bold = true;
                    sheet.Cells[2, 1].Style.Fill.PatternType = OfficeOpenXml.Style.ExcelFillStyle.Solid;
                    sheet.Cells[2, 1].Style.Fill.BackgroundColor.SetColor(Color.LightGray);

                    sheetAlternate.Cells[2, 1].Value = "Module-Slot";
                    sheetAlternate.Cells[2, 1].Style.Font.Bold = true;
                    sheetAlternate.Cells[2, 1].Style.Fill.PatternType = OfficeOpenXml.Style.ExcelFillStyle.Solid;
                    sheetAlternate.Cells[2, 1].Style.Fill.BackgroundColor.SetColor(Color.LightGray);

                    for (int i = 0; i < jobNames.Count; i++)
                    {
                        // Row 1: JobName
                        sheet.Cells[1, i + 2].Value = jobNames[i];
                        sheet.Cells[1, i + 2].Style.Font.Bold = true;
                        sheet.Cells[1, i + 2].Style.Fill.PatternType = OfficeOpenXml.Style.ExcelFillStyle.Solid;
                        sheet.Cells[1, i + 2].Style.Fill.BackgroundColor.SetColor(Color.LightBlue);
                        sheet.Cells[1, i + 2].Style.HorizontalAlignment = OfficeOpenXml.Style.ExcelHorizontalAlignment.Center;

                        sheetAlternate.Cells[1, i + 2].Value = jobNames[i];
                        sheetAlternate.Cells[1, i + 2].Style.Font.Bold = true;
                        sheetAlternate.Cells[1, i + 2].Style.Fill.PatternType = OfficeOpenXml.Style.ExcelFillStyle.Solid;
                        sheetAlternate.Cells[1, i + 2].Style.Fill.BackgroundColor.SetColor(Color.LightBlue);
                        sheetAlternate.Cells[1, i + 2].Style.HorizontalAlignment = OfficeOpenXml.Style.ExcelHorizontalAlignment.Center;

                        // Row 2: "PartNumber"
                        sheet.Cells[2, i + 2].Value = "PartNumber";
                        sheet.Cells[2, i + 2].Style.Font.Bold = true;
                        sheet.Cells[2, i + 2].Style.Fill.PatternType = OfficeOpenXml.Style.ExcelFillStyle.Solid;
                        sheet.Cells[2, i + 2].Style.Fill.BackgroundColor.SetColor(Color.LightGray);

                        sheetAlternate.Cells[2, i + 2].Value = "PartNumber";
                        sheetAlternate.Cells[2, i + 2].Style.Font.Bold = true;
                        sheetAlternate.Cells[2, i + 2].Style.Fill.PatternType = OfficeOpenXml.Style.ExcelFillStyle.Solid;
                        sheetAlternate.Cells[2, i + 2].Style.Fill.BackgroundColor.SetColor(Color.LightGray);
                    }

                    sheet.Cells[2, jobNames.Count + 2].Value = "Common";
                    sheet.Cells[2, jobNames.Count + 2].Style.Font.Bold = true;
                    sheet.Cells[2, jobNames.Count + 2].Style.Fill.PatternType = OfficeOpenXml.Style.ExcelFillStyle.Solid;
                    sheet.Cells[2, jobNames.Count + 2].Style.Fill.BackgroundColor.SetColor(Color.LightGray);
                    sheet.Cells[2, jobNames.Count + 2].Style.HorizontalAlignment = OfficeOpenXml.Style.ExcelHorizontalAlignment.Center;

                    sheetAlternate.Cells[2, jobNames.Count + 2].Value = "Common";
                    sheetAlternate.Cells[2, jobNames.Count + 2].Style.Font.Bold = true;
                    sheetAlternate.Cells[2, jobNames.Count + 2].Style.Fill.PatternType = OfficeOpenXml.Style.ExcelFillStyle.Solid;
                    sheetAlternate.Cells[2, jobNames.Count + 2].Style.Fill.BackgroundColor.SetColor(Color.LightGray);
                    sheetAlternate.Cells[2, jobNames.Count + 2].Style.HorizontalAlignment = OfficeOpenXml.Style.ExcelHorizontalAlignment.Center;

                    // Write data rows starting at row 3
                    int rowNum = 3;
                    int altRowNum = 3;
                    var mainModuleRanges = new List<(int StartRow, int EndRow)>();
                    var altModuleRanges = new List<(int StartRow, int EndRow)>();
                    string? currentMainModule = null;
                    int currentMainStart = 3;
                    string? currentAltModule = null;
                    int currentAltStart = 3;
                    
                    foreach (var moduleSlot in sortedModuleSlots)
                    {
                        if (cancelRequested)
                            break;

                        var moduleKey = GetModuleKey(moduleSlot);
                        if (currentMainModule == null)
                        {
                            currentMainModule = moduleKey;
                            currentMainStart = rowNum;
                        }
                        else if (!string.Equals(currentMainModule, moduleKey, StringComparison.OrdinalIgnoreCase))
                        {
                            mainModuleRanges.Add((currentMainStart, rowNum - 1));
                            rowNum += 2; // 2 filas en blanco antes del nuevo encabezado de módulo
                            WriteModuleDivisionHeader(sheet, rowNum, moduleKey, jobNames);
                            rowNum += 2; // encabezado de 2 filas
                            currentMainModule = moduleKey;
                            currentMainStart = rowNum;
                        }

                        sheet.Cells[rowNum, 1].Value = moduleSlot;

                        var partsPerJob = new Dictionary<string, List<string>>();
                        var hasAnyAlternate = false;

                        for (int j = 0; j < jobNames.Count; j++)
                        {
                            var jobName = jobNames[j];
                            var parts = dataByJobName[jobName]
                                .Where(x => x.ModuleSlot == moduleSlot)
                                .SelectMany(x => x.PartNumbers)
                                .Where(p => !string.IsNullOrWhiteSpace(p))
                                .Distinct()
                                .OrderBy(x => x)
                                .ToList();

                            partsPerJob[jobName] = parts;

                            // Si hay más de 1 PartNumber, usar el primero en main, los demás a Alternos
                            if (parts.Count > 0)
                            {
                                sheet.Cells[rowNum, j + 2].Value = parts[0];
                                
                                if (parts.Count > 1)
                                    hasAnyAlternate = true;
                            }
                            else
                            {
                                sheet.Cells[rowNum, j + 2].Value = "";
                            }
                        }

                        // Calculate Common for main sheet (usando solo el primer PartNumber de cada job)
                        var mainPartsPerJob = new Dictionary<string, List<string>>();
                        foreach (var kvp in partsPerJob)
                        {
                            mainPartsPerJob[kvp.Key] = kvp.Value.Count > 0 ? new List<string> { kvp.Value[0] } : new List<string>();
                        }

                        var (highlightJobs, comiunColor, comiunValue) = CalculateCellColors(mainPartsPerJob, jobNames);

                        // Apply yellow only to the cells that differ (not the whole row)
                        for (int j = 0; j < jobNames.Count; j++)
                        {
                            var jobName = jobNames[j];
                            if (!highlightJobs.Contains(jobName))
                                continue;

                            if (!mainPartsPerJob.TryGetValue(jobName, out var partsInJob) || partsInJob.Count == 0)
                                continue;

                            var cell = sheet.Cells[rowNum, j + 2];
                            cell.Style.Fill.PatternType = OfficeOpenXml.Style.ExcelFillStyle.Solid;
                            cell.Style.Fill.BackgroundColor.SetColor(Color.Yellow);
                            cell.Style.Font.Color.SetColor(Color.Black);
                        }

                        // Apply Common color and value
                        var comiunCell = sheet.Cells[rowNum, jobNames.Count + 2];
                        comiunCell.Value = comiunValue;
                        if (comiunColor.HasValue)
                        {
                            comiunCell.Style.Fill.PatternType = OfficeOpenXml.Style.ExcelFillStyle.Solid;
                            comiunCell.Style.Fill.BackgroundColor.SetColor(comiunColor.Value);
                            comiunCell.Style.Font.Color.SetColor(Color.White);
                        }

                        rowNum++;

                        // Si hay alternates, agregar filas en la hoja Alternos
                        if (hasAnyAlternate)
                        {
                            if (currentAltModule == null)
                            {
                                currentAltModule = moduleKey;
                                currentAltStart = altRowNum;
                            }
                            else if (!string.Equals(currentAltModule, moduleKey, StringComparison.OrdinalIgnoreCase))
                            {
                                altModuleRanges.Add((currentAltStart, altRowNum - 1));
                                altRowNum += 2; // 2 filas en blanco antes del nuevo encabezado de módulo
                                WriteModuleDivisionHeader(sheetAlternate, altRowNum, moduleKey, jobNames);
                                altRowNum += 2; // encabezado de 2 filas
                                currentAltModule = moduleKey;
                                currentAltStart = altRowNum;
                            }

                            // Para cada PartNumber alternativo (índice 1+)
                            int maxAlternates = partsPerJob.Values.Max(list => list.Count);
                            
                            for (int altIdx = 1; altIdx < maxAlternates; altIdx++)
                            {
                                sheetAlternate.Cells[altRowNum, 1].Value = moduleSlot;

                                var altPartsPerJob = new Dictionary<string, List<string>>();

                                for (int j = 0; j < jobNames.Count; j++)
                                {
                                    var jobName = jobNames[j];
                                    var parts = partsPerJob[jobName];

                                    if (altIdx < parts.Count)
                                    {
                                        sheetAlternate.Cells[altRowNum, j + 2].Value = parts[altIdx];
                                        altPartsPerJob[jobName] = new List<string> { parts[altIdx] };
                                    }
                                    else
                                    {
                                        sheetAlternate.Cells[altRowNum, j + 2].Value = "";
                                        altPartsPerJob[jobName] = new List<string>();
                                    }
                                }

                                // Calculate Common for alternate row
                                var (altHighlightJobs, altComiunColor, altComiunValue) = CalculateCellColors(altPartsPerJob, jobNames);

                                for (int j = 0; j < jobNames.Count; j++)
                                {
                                    var jobName = jobNames[j];
                                    if (!altHighlightJobs.Contains(jobName))
                                        continue;

                                    if (!altPartsPerJob.TryGetValue(jobName, out var partsInJob) || partsInJob.Count == 0)
                                        continue;

                                    var cell = sheetAlternate.Cells[altRowNum, j + 2];
                                    cell.Style.Fill.PatternType = OfficeOpenXml.Style.ExcelFillStyle.Solid;
                                    cell.Style.Fill.BackgroundColor.SetColor(Color.Yellow);
                                    cell.Style.Font.Color.SetColor(Color.Black);
                                }

                                var altComiunCell = sheetAlternate.Cells[altRowNum, jobNames.Count + 2];
                                altComiunCell.Value = altComiunValue;
                                if (altComiunColor.HasValue)
                                {
                                    altComiunCell.Style.Fill.PatternType = OfficeOpenXml.Style.ExcelFillStyle.Solid;
                                    altComiunCell.Style.Fill.BackgroundColor.SetColor(altComiunColor.Value);
                                    altComiunCell.Style.Font.Color.SetColor(Color.White);
                                }

                                altRowNum++;
                            }
                        }
                    }

                    if (currentMainModule != null && rowNum > currentMainStart)
                        mainModuleRanges.Add((currentMainStart, rowNum - 1));

                    if (currentAltModule != null && altRowNum > currentAltStart)
                        altModuleRanges.Add((currentAltStart, altRowNum - 1));

                    // Set column widths
                    sheet.Column(1).Width = 15;
                    sheetAlternate.Column(1).Width = 15;
                    for (int i = 2; i <= jobNames.Count + 2; i++)
                    {
                        sheet.Column(i).Width = 20;
                        sheetAlternate.Column(i).Width = 20;
                    }

                    if (rowNum > 3)
                        ApplyTableBordersByModule(sheet, 1, jobNames.Count + 2, mainModuleRanges);

                    if (altRowNum > 3)
                        ApplyTableBordersByModule(sheetAlternate, 1, jobNames.Count + 2, altModuleRanges);

                    if (cancelRequested)
                        return;

                    package.SaveAs(new System.IO.FileInfo(sfd.FileName));
                }

                if (!loadingForm.IsDisposed)
                    loadingForm.Close();
                MessageBox.Show("File exported successfully.", "Success", MessageBoxButtons.OK, MessageBoxIcon.Information);
            }
            catch (Exception ex)
            {
                if (!loadingForm.IsDisposed)
                    loadingForm.Close();
                MessageBox.Show($"Export error: {ex.Message}", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        private List<string> SortModuleSlots(List<string> moduleSlots)
        {
            return moduleSlots
                .OrderBy(x =>
                {
                    var parts = x.Split('-');
                    if (parts.Length == 2 && int.TryParse(parts[0], out int module) && int.TryParse(parts[1], out int slot))
                    {
                        return (module, slot);
                    }
                    return (int.MaxValue, int.MaxValue);
                })
                .ThenBy(x => x)
                .ToList();
        }

        private static string BuildModuleSlotWithoutSide(string moduleNumber, string location)
        {
            var module = (moduleNumber ?? string.Empty).Trim();
            var slot = (location ?? string.Empty).Trim();

            if (!string.IsNullOrWhiteSpace(module) && !string.IsNullOrWhiteSpace(slot))
                return $"{module}-{slot}";

            if (!string.IsNullOrWhiteSpace(module))
                return module;

            return slot;
        }

        private static string GetModuleKey(string moduleSlot)
        {
            var value = (moduleSlot ?? string.Empty).Trim();
            if (string.IsNullOrWhiteSpace(value))
                return string.Empty;

            var dash = value.IndexOf('-');
            return dash > 0 ? value.Substring(0, dash).Trim() : value;
        }

        private static void ApplyTableBordersByModule(
            OfficeOpenXml.ExcelWorksheet sheet,
            int firstCol,
            int lastCol,
            IEnumerable<(int StartRow, int EndRow)> moduleRanges)
        {
            foreach (var range in moduleRanges)
            {
                if (range.EndRow < range.StartRow)
                    continue;

                var block = sheet.Cells[range.StartRow, firstCol, range.EndRow, lastCol];
                block.Style.Border.Top.Style = OfficeOpenXml.Style.ExcelBorderStyle.Thin;
                block.Style.Border.Bottom.Style = OfficeOpenXml.Style.ExcelBorderStyle.Thin;
                block.Style.Border.Left.Style = OfficeOpenXml.Style.ExcelBorderStyle.Thin;
                block.Style.Border.Right.Style = OfficeOpenXml.Style.ExcelBorderStyle.Thin;
                block.Style.Border.Top.Color.SetColor(Color.Black);
                block.Style.Border.Bottom.Color.SetColor(Color.Black);
                block.Style.Border.Left.Color.SetColor(Color.Black);
                block.Style.Border.Right.Color.SetColor(Color.Black);

                var top = sheet.Cells[range.StartRow, firstCol, range.StartRow, lastCol];
                top.Style.Border.Top.Style = OfficeOpenXml.Style.ExcelBorderStyle.Medium;
                top.Style.Border.Top.Color.SetColor(Color.Black);

                var bottom = sheet.Cells[range.EndRow, firstCol, range.EndRow, lastCol];
                bottom.Style.Border.Bottom.Style = OfficeOpenXml.Style.ExcelBorderStyle.Medium;
                bottom.Style.Border.Bottom.Color.SetColor(Color.Black);

                var left = sheet.Cells[range.StartRow, firstCol, range.EndRow, firstCol];
                left.Style.Border.Left.Style = OfficeOpenXml.Style.ExcelBorderStyle.Medium;
                left.Style.Border.Left.Color.SetColor(Color.Black);

                var right = sheet.Cells[range.StartRow, lastCol, range.EndRow, lastCol];
                right.Style.Border.Right.Style = OfficeOpenXml.Style.ExcelBorderStyle.Medium;
                right.Style.Border.Right.Color.SetColor(Color.Black);
            }
        }

        private static void WriteModuleDivisionHeader(
            OfficeOpenXml.ExcelWorksheet sheet,
            int row,
            string moduleKey,
            List<string> jobNames)
        {
            var rowJobs = row;
            var rowLabels = row + 1;

            sheet.Cells[rowJobs, 1].Value = string.Empty;

            for (int j = 0; j < jobNames.Count; j++)
            {
                sheet.Cells[rowJobs, j + 2].Value = jobNames[j];
                sheet.Cells[rowJobs, j + 2].Style.Font.Bold = true;
                sheet.Cells[rowJobs, j + 2].Style.HorizontalAlignment = OfficeOpenXml.Style.ExcelHorizontalAlignment.Center;
                sheet.Cells[rowJobs, j + 2].Style.Fill.PatternType = OfficeOpenXml.Style.ExcelFillStyle.Solid;
                sheet.Cells[rowJobs, j + 2].Style.Fill.BackgroundColor.SetColor(Color.LightBlue);
            }

            sheet.Cells[rowJobs, jobNames.Count + 2].Value = string.Empty;

            var leftTopCell = sheet.Cells[rowJobs, 1];
            leftTopCell.Style.Fill.PatternType = OfficeOpenXml.Style.ExcelFillStyle.Solid;
            leftTopCell.Style.Fill.BackgroundColor.SetColor(Color.White);

            var rightTopCell = sheet.Cells[rowJobs, jobNames.Count + 2];
            rightTopCell.Style.Fill.PatternType = OfficeOpenXml.Style.ExcelFillStyle.Solid;
            rightTopCell.Style.Fill.BackgroundColor.SetColor(Color.White);

            sheet.Cells[rowLabels, 1].Value = "Module-Slot";
            sheet.Cells[rowLabels, 1].Style.Font.Bold = true;
            sheet.Cells[rowLabels, 1].Style.Fill.PatternType = OfficeOpenXml.Style.ExcelFillStyle.Solid;
            sheet.Cells[rowLabels, 1].Style.Fill.BackgroundColor.SetColor(Color.LightGray);

            for (int j = 0; j < jobNames.Count; j++)
            {
                sheet.Cells[rowLabels, j + 2].Value = "PartNumber";
                sheet.Cells[rowLabels, j + 2].Style.Font.Bold = true;
                sheet.Cells[rowLabels, j + 2].Style.Fill.PatternType = OfficeOpenXml.Style.ExcelFillStyle.Solid;
                sheet.Cells[rowLabels, j + 2].Style.Fill.BackgroundColor.SetColor(Color.LightGray);
            }

            sheet.Cells[rowLabels, jobNames.Count + 2].Value = "Common";
            sheet.Cells[rowLabels, jobNames.Count + 2].Style.Font.Bold = true;
            sheet.Cells[rowLabels, jobNames.Count + 2].Style.HorizontalAlignment = OfficeOpenXml.Style.ExcelHorizontalAlignment.Center;
            sheet.Cells[rowLabels, jobNames.Count + 2].Style.Fill.PatternType = OfficeOpenXml.Style.ExcelFillStyle.Solid;
            sheet.Cells[rowLabels, jobNames.Count + 2].Style.Fill.BackgroundColor.SetColor(Color.LightGray);

            var headerRange = sheet.Cells[rowJobs, 1, rowLabels, jobNames.Count + 2];
            headerRange.Style.Border.Top.Style = OfficeOpenXml.Style.ExcelBorderStyle.Thin;
            headerRange.Style.Border.Bottom.Style = OfficeOpenXml.Style.ExcelBorderStyle.Thin;
            headerRange.Style.Border.Left.Style = OfficeOpenXml.Style.ExcelBorderStyle.Thin;
            headerRange.Style.Border.Right.Style = OfficeOpenXml.Style.ExcelBorderStyle.Thin;
            headerRange.Style.Border.Top.Color.SetColor(Color.Black);
            headerRange.Style.Border.Bottom.Color.SetColor(Color.Black);
            headerRange.Style.Border.Left.Color.SetColor(Color.Black);
            headerRange.Style.Border.Right.Color.SetColor(Color.Black);
        }

        private static string GetUniqueName(string? baseName, HashSet<string> used)
        {
            var normalized = string.IsNullOrWhiteSpace(baseName) ? "Unknown" : baseName.Trim();
            if (used.Add(normalized))
                return normalized;

            var n = 2;
            while (true)
            {
                var candidate = $"{normalized} ({n})";
                if (used.Add(candidate))
                    return candidate;
                n++;
            }
        }

        private (HashSet<string> HighlightJobs, Color?, string) CalculateCellColors(Dictionary<string, List<string>> partsPerJob, List<string> jobNames)
        {
            // Regla para "Comiun": si existe al menos un PartNumber repetido entre 2+ JobName,
            // entonces Comiun = true (aunque otras celdas estén en blanco).

            // Normalizar: quitar blancos y trabajar con sets por job
            var jobPartSets = new Dictionary<string, HashSet<string>>(StringComparer.OrdinalIgnoreCase);
            foreach (var job in jobNames)
            {
                var set = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
                if (partsPerJob.TryGetValue(job, out var parts))
                {
                    foreach (var p in parts)
                    {
                        if (!string.IsNullOrWhiteSpace(p))
                            set.Add(p.Trim());
                    }
                }
                jobPartSets[job] = set;
            }

            var nonEmptySets = jobPartSets.Values.Where(s => s.Count > 0).ToList();

            bool allSame = false;
            if (nonEmptySets.Count > 0)
            {
                var first = nonEmptySets[0];
                allSame = nonEmptySets.Skip(1).All(s => s.SetEquals(first));
            }

            // Detectar si algún PartNumber aparece en 2 o más jobs
            var partJobCount = new Dictionary<string, int>(StringComparer.OrdinalIgnoreCase);
            foreach (var kvp in jobPartSets)
            {
                foreach (var part in kvp.Value)
                {
                    partJobCount.TryGetValue(part, out int count);
                    partJobCount[part] = count + 1;
                }
            }

            // Si no hay ningún número de parte en ninguna columna, dejar Common en blanco
            if (partJobCount.Count == 0)
            {
                return (new HashSet<string>(StringComparer.OrdinalIgnoreCase), null, string.Empty);
            }

            // Nueva regla: si hay diferencias (amarillo), Comiun debe ser false.
            // Es decir, Comiun=True solo cuando TODOS los jobs con PartNumber coinciden (o la fila está vacía).
            if (allSame)
            {
                return (new HashSet<string>(StringComparer.OrdinalIgnoreCase), Color.Green, "True");
            }

            // Diferencias => Comiun=false (rojo). Resaltar en amarillo los PartNumber distintos.
            // Regla: si hay más de 1 variante (2, 3, etc), se deben marcar las celdas con PartNumber.
            var highlightJobs = new HashSet<string>(StringComparer.OrdinalIgnoreCase);

            // Highlight all non-empty cells when there is more than one variant
            var keySet = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
            var keyCounts = new Dictionary<string, int>(StringComparer.OrdinalIgnoreCase);
            var jobKey = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase);
            foreach (var job in jobNames)
            {
                var set = jobPartSets[job];
                if (set.Count == 0)
                    continue;

                var key = string.Join("|", set.OrderBy(x => x));
                jobKey[job] = key;
                keySet.Add(key);
                keyCounts.TryGetValue(key, out int c);
                keyCounts[key] = c + 1;
            }

            if (keySet.Count > 1)
            {
                // Caso especial: solo 1 celda distinta y las demás iguales => resaltar solo la distinta
                if (keyCounts.Count == 2 && jobKey.Count == 2)
                {
                    // Solo hay 2 celdas con datos y son diferentes => resaltar ambas
                    foreach (var job in jobKey.Keys)
                        highlightJobs.Add(job);
                }
                else if (keyCounts.Count == 2 && keyCounts.Values.Min() == 1)
                {
                    var modeKey = keyCounts
                        .OrderByDescending(kvp => kvp.Value)
                        .ThenBy(kvp => kvp.Key, StringComparer.OrdinalIgnoreCase)
                        .First().Key;

                    foreach (var job in jobNames)
                    {
                        if (!jobKey.TryGetValue(job, out var key))
                            continue;
                        if (!string.Equals(key, modeKey, StringComparison.OrdinalIgnoreCase))
                            highlightJobs.Add(job);
                    }
                }
                else
                {
                    // En otros casos (2/3+ variantes), resaltar todas las celdas con PartNumber
                    foreach (var job in jobNames)
                    {
                        if (jobPartSets[job].Count > 0)
                            highlightJobs.Add(job);
                    }
                }
            }

            return (highlightJobs, Color.Red, "False");
        }

        private void EnableTabClosing(TabControl tc)
        {
            tc.DrawMode = TabDrawMode.OwnerDrawFixed;
            tc.Padding = new Point(18, 4);

            tc.DrawItem += (s, e) =>
            {
                var tabPage = tc.TabPages[e.Index];
                var tabRect = tc.GetTabRect(e.Index);

                bool isSelected = (e.State & DrawItemState.Selected) == DrawItemState.Selected;
                var backColor = GetTabBackColor();

                using var backBrush = new SolidBrush(backColor);
                e.Graphics.FillRectangle(backBrush, tabRect);

                if (isSelected)
                {
                    using var borderPen = new Pen(SystemColors.ControlDark, 1);
                    var r = tabRect;
                    r.Width -= 1;
                    r.Height -= 1;
                    e.Graphics.DrawRectangle(borderPen, r);
                }

                var textRect = new Rectangle(tabRect.X + 2, tabRect.Y + 4, tabRect.Width - 18, tabRect.Height - 4);
                TextRenderer.DrawText(e.Graphics, tabPage.Text, tc.Font, textRect, SystemColors.ControlText, TextFormatFlags.Left | TextFormatFlags.VerticalCenter);

                var closeRect = GetCloseRect(tabRect);
                using var pen = new Pen(Color.DarkRed, 2);
                e.Graphics.SmoothingMode = SmoothingMode.AntiAlias;
                e.Graphics.DrawLine(pen, closeRect.Left, closeRect.Top, closeRect.Right, closeRect.Bottom);
                e.Graphics.DrawLine(pen, closeRect.Right, closeRect.Top, closeRect.Left, closeRect.Bottom);
            };

            tc.MouseDown += (s, e) =>
            {
                for (int i = 0; i < tc.TabPages.Count; i++)
                {
                    var tabRect = tc.GetTabRect(i);
                    var closeRect = GetCloseRect(tabRect);
                    if (closeRect.Contains(e.Location))
                    {
                        var page = tc.TabPages[i];
                        tc.TabPages.RemoveAt(i);
                        page.Dispose();
                        break;
                    }
                }
            };

            static Rectangle GetCloseRect(Rectangle tabRect)
            {
                return new Rectangle(tabRect.Right - 14, tabRect.Top + 6, 10, 10);
            }

            static Color GetTabBackColor()
            {
                // Mismo color para todas las pestañas
                return SystemColors.Control;
            }
        }

        private static bool TryParseHeaders(string line, out string[] headers, out char delimiter)
        {
            char[] candidates = new[] { '\t', ',', ';' };
            foreach (var sep in candidates)
            {
                var cols = ParseDelimitedLine(line, sep)
                    .Select(NormalizeCell)
                    .ToArray();

                if (cols.Any(x => string.Equals(x, "ModuleNumber", StringComparison.OrdinalIgnoreCase)) &&
                    cols.Any(x => string.Equals(x, "PartNumber", StringComparison.OrdinalIgnoreCase)) &&
                    cols.Any(x => string.Equals(x, "Location", StringComparison.OrdinalIgnoreCase)))
                {
                    headers = cols;
                    delimiter = sep;
                    return true;
                }
            }

            headers = Array.Empty<string>();
            delimiter = ',';
            return false;
        }

        private static IEnumerable<string> ParseDelimitedLine(string line, char delimiter)
        {
            if (string.IsNullOrEmpty(line))
                yield break;

            var current = new System.Text.StringBuilder();
            bool inQuotes = false;
            char quoteChar = '\0';

            for (int i = 0; i < line.Length; i++)
            {
                char c = line[i];

                if (inQuotes)
                {
                    if (c == quoteChar)
                    {
                        // Handle escaped quotes by doubling
                        if (i + 1 < line.Length && line[i + 1] == quoteChar)
                        {
                            current.Append(c);
                            i++;
                        }
                        else
                        {
                            inQuotes = false;
                        }
                    }
                    else
                    {
                        current.Append(c);
                    }
                }
                else
                {
                    if (c == '"' || c == '\'')
                    {
                        inQuotes = true;
                        quoteChar = c;
                    }
                    else if (c == delimiter)
                    {
                        yield return current.ToString();
                        current.Clear();
                    }
                    else
                    {
                        current.Append(c);
                    }
                }
            }

            yield return current.ToString();
        }

        private static string NormalizeCell(string value)
        {
            return value.Trim().Trim('"', '\'');
        }

        private static string NormalizePath(string? value)
        {
            return (value ?? string.Empty).Trim().Trim('"');
        }

    }
}