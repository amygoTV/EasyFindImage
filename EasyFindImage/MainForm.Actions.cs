using System;
using System.Collections.Generic;
using System.Data;
using System.Data.OleDb;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Text;
using System.Windows.Forms;

namespace EasyFindImage
{
    public partial class MainForm
    {
        // ==================== Tab 1 事件 ====================

        private void BtnSummary_Click(object sender, EventArgs e)
        {
            string folder = txtSourceFolder.Text.Trim();
            if (!Directory.Exists(folder))
            {
                MessageBox.Show(this, "请选择有效的源文件夹", "提示", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return;
            }

            pgbStatus.Value = 0;
            lblStatus.Text = "正在汇总表格...";
            _summaryTable.Clear();
            dgvSummary.DataSource = null;

            try
            {
                var files = Directory.GetFiles(folder, "*.*", SearchOption.AllDirectories);
                int count = 0;
                foreach (var file in files)
                {
                    if (!IsTableFile(file)) continue;
                    count++;
                }

                int processed = 0;
                bool firstTable = true;

                foreach (var file in files)
                {
                    if (!IsTableFile(file)) continue;
                    processed++;
                    pgbStatus.Value = (int)((double)processed / count * 100);
                    lblStatus.Text = $"正在读取：{Path.GetFileName(file)}";
                    Application.DoEvents();

                    DataTable dt;
                    try
                    {
                        dt = ReadTableFile(file);
                    }
                    catch (Exception ex)
                    {
                        AppendLog($"跳过文件（读取失败）：{file}\n  错误：{ex.Message}");
                        continue;
                    }

                    if (dt.Rows.Count == 0) continue;

                    if (firstTable)
                    {
                        _summaryTable = dt.Clone();
                        firstTable = false;
                    }

                    foreach (DataRow row in dt.Rows)
                    {
                        // 按列名匹配复制数据，避免不同文件列数不一致导致“输入数组长度大于此表中的列数”错误
                        var values = new object[_summaryTable.Columns.Count];
                        for (int i = 0; i < _summaryTable.Columns.Count; i++)
                        {
                            string colName = _summaryTable.Columns[i].ColumnName;
                            if (dt.Columns.Contains(colName))
                                values[i] = row[colName];
                            else
                                values[i] = DBNull.Value;
                        }
                        _summaryTable.Rows.Add(values);
                    }
                }

                // 应用用户自定义数据清理条件（从条件列表读取，同一列多个关键字用逗号分隔表示"或"关系）

                // 1. 先自动移除残留的标题行（整行所有单元格值都等于对应列名的行）
                _summaryTable = RemoveHeaderRows(_summaryTable);

                // 2. 再应用用户自定义清理条件
                string cleanText = BuildCleanConditionsText();
                if (!string.IsNullOrEmpty(cleanText))
                    _summaryTable = ApplyCustomCleanConditions(_summaryTable, cleanText);

                // 3. 可选：只保留所有文件共有的列（去掉全空列）
                if (cbxKeepOnlyCommonCols != null && cbxKeepOnlyCommonCols.Checked)
                    _summaryTable = KeepOnlyCommonColumns(_summaryTable);

                if (_summaryTable.Rows.Count == 0)
                {
                    MessageBox.Show(this, "清理后未剩余任何数据行", "提示", MessageBoxButtons.OK, MessageBoxIcon.Information);
                    lblStatus.Text = "清理后无数据";
                    pgbStatus.Value = 100;
                    return;
                }

                _bindingSource.DataSource = _summaryTable;
                dgvSummary.DataSource = _bindingSource;
                lblStatus.Text = $"汇总完成，共 {_summaryTable.Rows.Count} 行数据";
                pgbStatus.Value = 100;

                // 设置列宽：根据内容自动调整，不强制填充，超出可视区域时显示横向滚动条
                foreach (DataGridViewColumn col in dgvSummary.Columns)
                {
                    col.AutoSizeMode = DataGridViewAutoSizeColumnMode.AllCells;
                    if (col.Width < 80) col.Width = 80; // 最小宽度
                }

                // 刷新列名下拉框
                RefreshColumnComboBoxes();

                if (_summaryTable.Rows.Count == 0)
                    MessageBox.Show(this, "未在所选文件夹中发现可读取的表格数据", "提示", MessageBoxButtons.OK, MessageBoxIcon.Information);
            }
            catch (Exception ex)
            {
                MessageBox.Show(this, $"汇总出错：{ex.Message}", "错误", MessageBoxButtons.OK, MessageBoxIcon.Error);
                lblStatus.Text = "汇总失败";
            }
        }

        private void BtnFilter_Click(object sender, EventArgs e)
        {
            if (_summaryTable == null || _summaryTable.Rows.Count == 0)
            {
                MessageBox.Show(this, "没有可筛选的数据，请先汇总表格", "提示", MessageBoxButtons.OK, MessageBoxIcon.Information);
                return;
            }

            string column = cbxFilterColumn.Text.Trim();
            string value = txtFilterValue.Text.Trim();

            if (string.IsNullOrEmpty(column))
            {
                MessageBox.Show(this, "请选择列名", "提示", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return;
            }

            if (!_summaryTable.Columns.Contains(column))
            {
                MessageBox.Show(this, $"列“{column}”不存在于汇总表中", "提示", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return;
            }

            try
            {
                string filter = $"[{column}] LIKE '%{value.Replace("'", "''")}%'";
                _bindingSource.Filter = filter;
                lblStatus.Text = $"筛选后显示 {_bindingSource.Count} 行";
            }
            catch (Exception ex)
            {
                MessageBox.Show(this, $"筛选表达式错误：{ex.Message}", "错误", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        private void BtnClearFilter_Click(object sender, EventArgs e)
        {
            cbxFilterColumn.SelectedIndex = -1;
            txtFilterValue.Clear();
            _bindingSource.RemoveFilter();
            lblStatus.Text = $"已清除筛选，显示 {_summaryTable.Rows.Count} 行";
        }

        private void BtnExport_Click(object sender, EventArgs e)
        {
            if (_summaryTable == null || _summaryTable.Rows.Count == 0)
            {
                MessageBox.Show(this, "没有数据可导出", "提示", MessageBoxButtons.OK, MessageBoxIcon.Information);
                return;
            }

            using (var dlg = new SaveFileDialog { Filter = "CSV 文件(*.csv)|*.csv", Title = "导出汇总结果", FileName = "汇总结果.csv" })
            {
                if (dlg.ShowDialog(this) == DialogResult.OK)
                {
                    try
                    {
                        var sb = new StringBuilder();
                        // 表头
                        for (int i = 0; i < _summaryTable.Columns.Count; i++)
                        {
                            if (i > 0) sb.Append(",");
                            sb.Append($"\"{_summaryTable.Columns[i].ColumnName}\"");
                        }
                        sb.AppendLine();

                        // 数据（尊重当前筛选后的视图）
                        var view = _bindingSource.List as DataView;
                        int rows = view != null ? view.Count : _summaryTable.Rows.Count;
                        for (int r = 0; r < rows; r++)
                        {
                            DataRow row = view != null ? view[r].Row : _summaryTable.Rows[r];
                            for (int c = 0; c < _summaryTable.Columns.Count; c++)
                            {
                                if (c > 0) sb.Append(",");
                                string val = row[c]?.ToString() ?? "";
                                if (val.Contains(",") || val.Contains("\"") || val.Contains("\n"))
                                    val = "\"" + val.Replace("\"", "\"\"") + "\"";
                                sb.Append(val);
                            }
                            sb.AppendLine();
                        }

                        File.WriteAllText(dlg.FileName, sb.ToString(), Encoding.UTF8);
                        MessageBox.Show(this, $"导出成功：{dlg.FileName}", "完成", MessageBoxButtons.OK, MessageBoxIcon.Information);
                    }
                    catch (Exception ex)
                    {
                        MessageBox.Show(this, $"导出失败：{ex.Message}", "错误", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    }
                }
            }
        }

        // ==================== 列名下拉框刷新 ====================

        /// <summary>
        /// 刷新筛选和清理的列名下拉框，从汇总表表头自动生成
        /// </summary>
        private void RefreshColumnComboBoxes()
        {
            cbxFilterColumn.Items.Clear();
            cbxCleanColumn.Items.Clear();

            if (_summaryTable != null)
            {
                foreach (DataColumn col in _summaryTable.Columns)
                {
                    cbxFilterColumn.Items.Add(col.ColumnName);
                    cbxCleanColumn.Items.Add(col.ColumnName);
                }
            }
        }

        // ==================== 数据清理条件管理 ====================

        private void BtnAddCleanCondition_Click(object sender, EventArgs e)
        {
            if (_summaryTable == null || _summaryTable.Rows.Count == 0)
            {
                MessageBox.Show(this, "没有数据，请先汇总表格", "提示", MessageBoxButtons.OK, MessageBoxIcon.Information);
                return;
            }

            string column = cbxCleanColumn.Text.Trim();
            string keywords = txtCleanKeywords.Text.Trim();

            if (string.IsNullOrEmpty(column))
            {
                MessageBox.Show(this, "请选择列名", "提示", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return;
            }

            // 允许关键字为空：表示删除该列为空白的行
            string condition = $"{column}:{keywords}";
            lstCleanConditions.Items.Add(condition);
            txtCleanKeywords.Clear();

            string tip = string.IsNullOrEmpty(keywords) ? "（该列为空白）" : "";
            lblStatus.Text = $"已添加清理条件：{condition}{tip}，共 {lstCleanConditions.Items.Count} 个";
        }

        private void BtnRemoveCleanCondition_Click(object sender, EventArgs e)
        {
            if (lstCleanConditions.SelectedIndex >= 0)
            {
                string removed = lstCleanConditions.SelectedItem?.ToString() ?? "";
                lstCleanConditions.Items.RemoveAt(lstCleanConditions.SelectedIndex);
                lblStatus.Text = $"已移除清理条件：{removed}，共 {lstCleanConditions.Items.Count} 个";
            }
        }

        private void BtnClearCleanConditions_Click(object sender, EventArgs e)
        {
            lstCleanConditions.Items.Clear();
            lblStatus.Text = "已清空所有清理条件";
        }

        private void BtnExecuteClean_Click(object sender, EventArgs e)
        {
            if (_summaryTable == null || _summaryTable.Rows.Count == 0)
            {
                MessageBox.Show(this, "没有数据，请先汇总表格", "提示", MessageBoxButtons.OK, MessageBoxIcon.Information);
                return;
            }

            // 至少有一个清理操作才执行
            if (lstCleanConditions.Items.Count == 0 && (cbxRemoveDuplicateSN == null || !cbxRemoveDuplicateSN.Checked))
            {
                MessageBox.Show(this, "请添加清理条件，或勾选“删除重复SN行”", "提示", MessageBoxButtons.OK, MessageBoxIcon.Information);
                return;
            }

            int beforeCount = _summaryTable.Rows.Count;

            // 1. 先自动移除残留的标题行
            _summaryTable = RemoveHeaderRows(_summaryTable);

            // 2. 应用清理条件（可选）
            string cleanText = BuildCleanConditionsText();
            if (!string.IsNullOrEmpty(cleanText))
                _summaryTable = ApplyCustomCleanConditions(_summaryTable, cleanText);

            // 3. 删除重复SN行（可选，保留第一个）
            int removedByDupSN = 0;
            if (cbxRemoveDuplicateSN != null && cbxRemoveDuplicateSN.Checked)
            {
                _summaryTable = RemoveDuplicateSnRows(_summaryTable, out removedByDupSN);
            }

            int afterCount = _summaryTable.Rows.Count;
            int removedCount = beforeCount - afterCount;

            // 刷新数据视图
            _bindingSource.DataSource = _summaryTable;
            dgvSummary.DataSource = _bindingSource;
            _bindingSource.RemoveFilter();

            string dupMsg = removedByDupSN > 0 ? $"（其中重复SN删除 {removedByDupSN} 行）" : "";
            lblStatus.Text = $"清理完成：删除 {removedCount} 行，剩余 {afterCount} 行 {dupMsg}";
            MessageBox.Show(this, $"执行清理完成！\n删除行数：{removedCount}\n剩余行数：{afterCount}\n{dupMsg}", "完成", MessageBoxButtons.OK, MessageBoxIcon.Information);
        }

        /// <summary>
        /// 从条件列表构建清理条件文本，格式：列名:关键字1,关键字2,列名2:关键字3
        /// 同一列多个关键字用逗号分隔，表示"或"关系（任一匹配即删除）
        /// </summary>
        private string BuildCleanConditionsText()
        {
            var parts = new List<string>();
            foreach (string item in lstCleanConditions.Items)
            {
                parts.Add(item);
            }
            return string.Join(",", parts);
        }

        /// <summary>
        /// 删除重复SN行（保留第一次出现的SN，删除后面重复的）
        /// </summary>
        private DataTable RemoveDuplicateSnRows(DataTable table, out int removedCount)
        {
            removedCount = 0;
            if (table.Rows.Count == 0) return table;

            // 查找SN列
            string snCol = null;
            foreach (DataColumn col in table.Columns)
            {
                if (col.ColumnName.Equals("SN", StringComparison.OrdinalIgnoreCase))
                {
                    snCol = col.ColumnName;
                    break;
                }
            }

            if (snCol == null) return table;

            var seenSns = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
            var rowsToRemove = new List<DataRow>();

            foreach (DataRow row in table.Rows)
            {
                if (row.RowState == DataRowState.Deleted) continue;

                string snVal = row[snCol]?.ToString()?.Trim() ?? "";
                if (string.IsNullOrEmpty(snVal)) continue;

                if (seenSns.Contains(snVal))
                {
                    rowsToRemove.Add(row);
                    removedCount++;
                }
                else
                {
                    seenSns.Add(snVal);
                }
            }

            foreach (var r in rowsToRemove) r.Delete();
            table.AcceptChanges();
            return table;
        }

        // ==================== Tab 2 事件 ====================

        private void BtnStartCopy_Click(object sender, EventArgs e)
        {
            string snFile = txtSnSource.Text.Trim();
            string imageRoot = txtImageRoot.Text.Trim();
            string targetDir = txtTargetDir.Text.Trim();
            bool searchAllDrives = rdoAllDrives.Checked;

            if (string.IsNullOrEmpty(snFile) || !File.Exists(snFile))
            {
                MessageBox.Show(this, "请选择有效的 SN 来源文件", "提示", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return;
            }
            if (!searchAllDrives && (string.IsNullOrEmpty(imageRoot) || !Directory.Exists(imageRoot)))
            {
                MessageBox.Show(this, "请选择有效的图片根目录", "提示", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return;
            }
            if (string.IsNullOrEmpty(targetDir))
            {
                MessageBox.Show(this, "请选择复制目标路径", "提示", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return;
            }

            btnStartCopy.Enabled = false;
            pgbCopy.Value = 0;
            lblCopyStatus.Text = "正在提取 SN...";
            lstLog.Items.Add("========== 开始任务 ==========");

            try
            {
                var sns = ExtractSns(snFile, rdoFromExcel.Checked);
                if (sns.Count == 0)
                {
                    AppendLog("未提取到任何 SN");
                    lblCopyStatus.Text = "未找到 SN";
                    return;
                }

                // 准备查找根目录列表
                List<string> searchRoots = new List<string>();
                if (searchAllDrives)
                {
                    AppendLog("全盘递归查找模式：正在扫描所有磁盘根目录...");
                    foreach (var drive in System.IO.DriveInfo.GetDrives())
                    {
                        if (drive.IsReady && drive.DriveType == System.IO.DriveType.Fixed)
                        {
                            searchRoots.Add(drive.RootDirectory.FullName);
                            AppendLog($"  发现磁盘：{drive.RootDirectory.FullName} ({drive.VolumeLabel})");
                        }
                    }
                    if (searchRoots.Count == 0)
                    {
                        MessageBox.Show(this, "未发现可用的固定磁盘", "提示", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                        return;
                    }
                }
                else
                {
                    searchRoots.Add(imageRoot);
                }

                AppendLog($"共提取到 {sns.Count} 个 SN，开始查找并复制（搜索根目录：{searchRoots.Count} 个）...");

                Directory.CreateDirectory(targetDir);

                int success = 0, fail = 0;
                for (int i = 0; i < sns.Count; i++)
                {
                    string sn = sns[i];
                    lblCopyStatus.Text = $"处理中 [{i + 1}/{sns.Count}]：{sn} | 已成功 {success}，失败 {fail}";
                    pgbCopy.Value = (int)((double)(i + 1) / sns.Count * 100);
                    Application.DoEvents();

                    string foundDir = null;
                    var allFoundDirs = new List<string>();
                    foreach (var root in searchRoots)
                    {
                        var foundInRoot = FindAllSnFoldersRecursive(root, sn);
                        allFoundDirs.AddRange(foundInRoot);
                    }

                    if (allFoundDirs.Count > 0)
                    {
                        // 按创建时间排序，取最早的一个
                        foundDir = allFoundDirs
                            .OrderBy(d => Directory.GetCreationTime(d))
                            .First();
                        if (allFoundDirs.Count > 1)
                        {
                            AppendLog($"[多匹配] {sn} 找到 {allFoundDirs.Count} 个文件夹，已选择创建时间最早的：{Path.GetFileName(foundDir)}");
                        }
                    }
                    if (foundDir == null)
                    {
                        AppendLog($"[未找到] {sn}");
                        fail++;
                        continue;
                    }

                    string dest = Path.Combine(targetDir, Path.GetFileName(foundDir));
                    if (Directory.Exists(dest))
                    {
                        dest = Path.Combine(targetDir, $"{Path.GetFileName(foundDir)}_{sn}");
                    }

                    try
                    {
                        CopyDirectory(foundDir, dest);
                        AppendLog($"[成功] {sn} -> {dest}");
                        success++;
                    }
                    catch (Exception ex)
                    {
                        AppendLog($"[失败] {sn}: {ex.Message}");
                        fail++;
                    }
                }

                lblCopyStatus.Text = $"完成：成功 {success} 个，失败 {fail} 个";
                AppendLog($"========== 任务结束：成功 {success}，失败 {fail} ==========");
                MessageBox.Show(this, $"复制完成！\n成功：{success}\n失败：{fail}", "完成", MessageBoxButtons.OK, MessageBoxIcon.Information);
            }
            catch (Exception ex)
            {
                MessageBox.Show(this, $"操作出错：{ex.Message}", "错误", MessageBoxButtons.OK, MessageBoxIcon.Error);
                AppendLog($"异常：{ex.Message}");
            }
            finally
            {
                btnStartCopy.Enabled = true;
                pgbCopy.Value = 100;
            }
        }

        // ==================== 业务方法 ====================

        private bool IsTableFile(string path)
        {
            string ext = Path.GetExtension(path).ToLowerInvariant();
            return ext == ".csv" || ext == ".xls" || ext == ".xlsx";
        }

        private DataTable ReadTableFile(string path)
        {
            string ext = Path.GetExtension(path).ToLowerInvariant();
            if (ext == ".csv")
                return ReadCsvFile(path);
            return ReadExcelFile(path);
        }

        private DataTable ReadCsvFile(string path)
        {
            var dt = new DataTable();
            using (var sr = new StreamReader(path, Encoding.Default))
            {
                string header = sr.ReadLine();
                if (header == null) return dt;

                string[] cols = header.Split(',');
                foreach (var c in cols)
                    dt.Columns.Add(c.Trim());

                while (!sr.EndOfStream)
                {
                    string line = sr.ReadLine();
                    if (string.IsNullOrWhiteSpace(line)) continue;
                    string[] vals = line.Split(',');
                    DataRow row = dt.NewRow();
                    for (int i = 0; i < vals.Length && i < dt.Columns.Count; i++)
                        row[i] = vals[i].Trim();
                    dt.Rows.Add(row);
                }
            }
            return dt;
        }

        private DataTable ReadExcelFile(string path)
        {
            string connStr;
            if (path.EndsWith(".xlsx", StringComparison.OrdinalIgnoreCase))
                connStr = $"Provider=Microsoft.ACE.OLEDB.12.0;Data Source={path};Extended Properties='Excel 12.0 Xml;HDR=YES;'";
            else
                connStr = $"Provider=Microsoft.Jet.OLEDB.4.0;Data Source={path};Extended Properties='Excel 8.0;HDR=YES;'";

            using (var conn = new OleDbConnection(connStr))
            {
                conn.Open();
                var schema = conn.GetOleDbSchemaTable(OleDbSchemaGuid.Tables, null);
                if (schema == null || schema.Rows.Count == 0)
                    throw new Exception("Excel 文件中未发现工作表");

                string sheet = schema.Rows[0]["TABLE_NAME"].ToString();
                using (var cmd = new OleDbCommand($"SELECT * FROM [{sheet}]", conn))
                using (var adapter = new OleDbDataAdapter(cmd))
                {
                    var dt = new DataTable();
                    adapter.Fill(dt);
                    return dt;
                }
            }
        }

        private System.Collections.Generic.List<string> ExtractSns(string filePath, bool fromExcel)
        {
            var sns = new System.Collections.Generic.List<string>();

            if (fromExcel)
            {
                DataTable dt = ReadTableFile(filePath);
                // 尝试查找名为 SN 或 sn 的列
                string snCol = null;
                foreach (DataColumn col in dt.Columns)
                {
                    if (col.ColumnName.Equals("SN", StringComparison.OrdinalIgnoreCase) ||
                        col.ColumnName.Equals("sn", StringComparison.OrdinalIgnoreCase))
                    {
                        snCol = col.ColumnName;
                        break;
                    }
                }

                if (snCol == null)
                {
                    // 若无 SN 列，使用第一列
                    snCol = dt.Columns.Count > 0 ? dt.Columns[0].ColumnName : null;
                }

                if (snCol != null)
                {
                    foreach (DataRow row in dt.Rows)
                    {
                        string sn = row[snCol]?.ToString().Trim();
                        if (!string.IsNullOrEmpty(sn) && !sns.Contains(sn))
                            sns.Add(sn);
                    }
                }
            }
            else
            {
                // TXT：每行一个 SN
                using (var sr = new StreamReader(filePath, Encoding.Default))
                {
                    while (!sr.EndOfStream)
                    {
                        string line = sr.ReadLine();
                        if (string.IsNullOrWhiteSpace(line)) continue;
                        string sn = line.Trim();
                        if (!sns.Contains(sn))
                            sns.Add(sn);
                    }
                }
            }

            return sns;
        }

        private string FindSnFolder(string root, string sn)
        {
            if (!Directory.Exists(root)) return null;

            // 1. 精确匹配文件夹名
            foreach (var dir in Directory.GetDirectories(root))
            {
                if (Path.GetFileName(dir).Equals(sn, StringComparison.OrdinalIgnoreCase))
                    return dir;
            }

            // 2. 模糊匹配：文件夹名包含 SN
            foreach (var dir in Directory.GetDirectories(root))
            {
                if (Path.GetFileName(dir).IndexOf(sn, StringComparison.OrdinalIgnoreCase) >= 0)
                    return dir;
            }

            return null;
        }

        /// <summary>
        /// 递归查找 SN 文件夹（从指定根目录开始全盘/子目录递归搜索）
        /// 先精确匹配，再模糊匹配，找到即返回
        /// </summary>
        private string FindSnFolderRecursive(string root, string sn)
        {
            if (!Directory.Exists(root)) return null;

            try
            {
                // 1. 当前目录下精确匹配文件夹名
                foreach (var dir in Directory.GetDirectories(root))
                {
                    if (Path.GetFileName(dir).Equals(sn, StringComparison.OrdinalIgnoreCase))
                        return dir;
                }

                // 2. 当前目录下模糊匹配：文件夹名包含 SN
                foreach (var dir in Directory.GetDirectories(root))
                {
                    if (Path.GetFileName(dir).IndexOf(sn, StringComparison.OrdinalIgnoreCase) >= 0)
                        return dir;
                }

                // 3. 递归搜索子目录
                foreach (var dir in Directory.GetDirectories(root))
                {
                    try
                    {
                        string result = FindSnFolderRecursive(dir, sn);
                        if (result != null) return result;
                    }
                    catch
                    {
                        // 忽略访问被拒绝的目录（如系统保护文件夹）
                    }
                }
            }
            catch
            {
                // 忽略根目录访问错误
            }

            return null;
        }

        /// <summary>
        /// 递归查找所有匹配的 SN 文件夹（返回完整路径列表）
        /// </summary>
        private List<string> FindAllSnFoldersRecursive(string root, string sn)
        {
            var results = new List<string>();
            if (!Directory.Exists(root)) return results;

            try
            {
                // 1. 当前目录下精确匹配文件夹名
                foreach (var dir in Directory.GetDirectories(root))
                {
                    if (Path.GetFileName(dir).Equals(sn, StringComparison.OrdinalIgnoreCase))
                        results.Add(dir);
                }

                // 2. 当前目录下模糊匹配：文件夹名包含 SN
                foreach (var dir in Directory.GetDirectories(root))
                {
                    if (Path.GetFileName(dir).IndexOf(sn, StringComparison.OrdinalIgnoreCase) >= 0)
                        results.Add(dir);
                }

                // 3. 递归搜索子目录
                foreach (var dir in Directory.GetDirectories(root))
                {
                    try
                    {
                        var subResults = FindAllSnFoldersRecursive(dir, sn);
                        results.AddRange(subResults);
                    }
                    catch
                    {
                        // 忽略访问被拒绝的目录
                    }
                }
            }
            catch
            {
                // 忽略根目录访问错误
            }

            return results;
        }

        private static void CopyDirectory(string sourceDir, string destDir)
        {
            Directory.CreateDirectory(destDir);

            foreach (var file in Directory.GetFiles(sourceDir))
            {
                string destFile = Path.Combine(destDir, Path.GetFileName(file));
                File.Copy(file, destFile, true);
            }

            foreach (var subDir in Directory.GetDirectories(sourceDir))
            {
                string destSub = Path.Combine(destDir, Path.GetFileName(subDir));
                CopyDirectory(subDir, destSub);
            }
        }

        private void AppendLog(string message)
        {
            lstLog.Items.Add($"[{DateTime.Now:HH:mm:ss}] {message}");
            lstLog.TopIndex = lstLog.Items.Count - 1;
            Application.DoEvents();
        }

        // ==================== 数据清理辅助方法 ====================

        private DataTable RemoveDuplicateRows(DataTable table)
        {
            if (table.Rows.Count == 0) return table;

            var seen = new HashSet<string>();
            var rowsToRemove = new List<DataRow>();

            foreach (DataRow row in table.Rows)
            {
                if (row.RowState == DataRowState.Deleted) continue;

                var sb = new System.Text.StringBuilder();
                foreach (object val in row.ItemArray)
                    sb.Append("|" + (val == DBNull.Value ? "" : val.ToString()));
                string key = sb.ToString();

                if (seen.Contains(key))
                    rowsToRemove.Add(row);
                else
                    seen.Add(key);
            }

            foreach (var r in rowsToRemove) r.Delete();
            table.AcceptChanges();
            return table;
        }

        private DataTable RemoveInvalidRows(DataTable table)
        {
            var rowsToRemove = new List<DataRow>();
            foreach (DataRow row in table.Rows)
            {
                if (row.RowState == DataRowState.Deleted) continue;

                bool allEmpty = true;
                foreach (object val in row.ItemArray)
                {
                    if (val != DBNull.Value && !string.IsNullOrWhiteSpace(val.ToString()))
                    {
                        allEmpty = false;
                        break;
                    }
                }
                if (allEmpty) rowsToRemove.Add(row);
            }

            foreach (var r in rowsToRemove) r.Delete();
            table.AcceptChanges();
            return table;
        }

        private DataTable RemoveBadSnRows(DataTable table)
        {
            // 找到 SN 列
            string snCol = null;
            foreach (DataColumn col in table.Columns)
            {
                if (col.ColumnName.Equals("SN", StringComparison.OrdinalIgnoreCase))
                {
                    snCol = col.ColumnName;
                    break;
                }
            }
            if (snCol == null) return table;

            var rowsToRemove = new List<DataRow>();
            foreach (DataRow row in table.Rows)
            {
                if (row.RowState == DataRowState.Deleted) continue;

                string val = row[snCol]?.ToString().Trim();
                if (string.IsNullOrEmpty(val)) continue;

                // 值等于 "SN"（不区分大小写）
                if (val.Equals("SN", StringComparison.OrdinalIgnoreCase))
                {
                    rowsToRemove.Add(row);
                    continue;
                }

                // 重复字母：所有字符都相同，且长度 >= 2（如 SSSSS、NNNNN、AAAAA）
                if (val.Length >= 2 && val.Distinct().Count() == 1)
                {
                    rowsToRemove.Add(row);
                }
            }

            foreach (var r in rowsToRemove) r.Delete();
            table.AcceptChanges();
            return table;
        }

        private DataTable RemoveRowsContaining999(DataTable table)
        {
            var rowsToRemove = new List<DataRow>();
            foreach (DataRow row in table.Rows)
            {
                if (row.RowState == DataRowState.Deleted) continue;

                bool remove = false;
                foreach (DataColumn col in table.Columns)
                {
                    string val = row[col]?.ToString() ?? "";
                    if (val.Contains("999") || val == "999.999")
                    {
                        remove = true;
                        break;
                    }
                }
                if (remove) rowsToRemove.Add(row);
            }

            foreach (var r in rowsToRemove) r.Delete();
            table.AcceptChanges();
            return table;
        }

        /// <summary>
        /// 移除残留的标题行：整行所有单元格值都等于对应列名的行（例如 SN 列的值恰好是 "SN"）
        /// </summary>
        private DataTable RemoveHeaderRows(DataTable table)
        {
            if (table.Rows.Count == 0) return table;

            var rowsToRemove = new List<DataRow>();
            foreach (DataRow row in table.Rows)
            {
                if (row.RowState == DataRowState.Deleted) continue;

                bool isHeaderRow = true;
                for (int i = 0; i < table.Columns.Count; i++)
                {
                    string cellVal = row[i]?.ToString()?.Trim() ?? "";
                    string colName = table.Columns[i].ColumnName?.Trim() ?? "";
                    if (!cellVal.Equals(colName, StringComparison.OrdinalIgnoreCase))
                    {
                        isHeaderRow = false;
                        break;
                    }
                }
                if (isHeaderRow) rowsToRemove.Add(row);
            }

            foreach (var r in rowsToRemove) r.Delete();
            table.AcceptChanges();
            return table;
        }

        /// <summary>
        /// 应用用户自定义清理条件。
        /// 格式：列名:关键字1,关键字2,列名2:关键字3
        /// 同一列多个关键字用逗号分隔，表示"或"关系（任一匹配即删除）。
        /// 例如：SN:999,888,状态:无效  表示删除 SN 列包含 999 或 888 的行，或状态列等于“无效”的行
        /// 空白条件示例：SN:  表示删除 SN 列为空白的行
        /// </summary>
        private DataTable ApplyCustomCleanConditions(DataTable table, string conditions)
        {
            if (table.Rows.Count == 0) return table;

            // 解析条件：按逗号分割，包含冒号的开始新列，不包含冒号的作为上一列的附加关键字（"或"关系）
            var conditionsByColumn = new Dictionary<string, List<string>>(StringComparer.OrdinalIgnoreCase);
            string currentColumn = null;

            var rawParts = conditions.Split(new[] { ',' }, StringSplitOptions.RemoveEmptyEntries);
            foreach (var part in rawParts)
            {
                var trimmed = part.Trim();
                if (string.IsNullOrEmpty(trimmed)) continue;

                var colonIndex = trimmed.IndexOf(':');
                if (colonIndex >= 0)
                {
                    // 新的列名:关键字
                    currentColumn = trimmed.Substring(0, colonIndex).Trim();
                    string keyword = trimmed.Substring(colonIndex + 1).Trim();
                    if (string.IsNullOrEmpty(currentColumn)) continue;

                    if (!conditionsByColumn.ContainsKey(currentColumn))
                        conditionsByColumn[currentColumn] = new List<string>();
                    // 允许空关键字：用于匹配该列为空白的行（如 SN: 表示SN列为空白）
                    conditionsByColumn[currentColumn].Add(keyword);
                }
                else
                {
                    // 附加关键字，归属于上一个列名（"或"关系）
                    // 如果前面没有列名，则忽略此部分（避免 NA, SN 等无列名条件被误解析）
                    if (currentColumn != null && conditionsByColumn.ContainsKey(currentColumn))
                    {
                        if (!string.IsNullOrEmpty(trimmed))
                            conditionsByColumn[currentColumn].Add(trimmed);
                    }
                }
            }

            if (conditionsByColumn.Count == 0) return table;

            var rowsToRemove = new List<DataRow>();

            foreach (DataRow row in table.Rows)
            {
                if (row.RowState == DataRowState.Deleted) continue;

                bool remove = false;
                foreach (var kvp in conditionsByColumn)
                {
                    string colName = kvp.Key;
                    var keywords = kvp.Value;

                    if (!table.Columns.Contains(colName)) continue;

                    // 鲁棒地获取单元格值：处理 null、DBNull.Value、空字符串
                    object rawVal = row[colName];
                    string cellVal = "";
                    if (rawVal != null && rawVal != DBNull.Value)
                        cellVal = rawVal.ToString().Trim();

                    foreach (string keyword in keywords)
                    {
                        if (string.IsNullOrEmpty(keyword))
                        {
                            // 空关键字：匹配该列为空白的行
                            if (string.IsNullOrWhiteSpace(cellVal))
                            {
                                remove = true;
                                break;
                            }
                        }
                        else
                        {
                            // 精确等于 或 包含
                            if (cellVal.Equals(keyword, StringComparison.OrdinalIgnoreCase) ||
                                cellVal.Contains(keyword))
                            {
                                remove = true;
                                break;
                            }
                        }
                    }
                    if (remove) break;
                }

                if (remove) rowsToRemove.Add(row);
            }

            foreach (var r in rowsToRemove) r.Delete();
            table.AcceptChanges();
            return table;
        }

        // ==================== 配置持久化（关闭保存，打开加载） ====================

        private void MainForm_Load(object sender, EventArgs e)
        {
            LoadConfigToUi();
        }

        private void MainForm_FormClosing(object sender, FormClosingEventArgs e)
        {
            SaveUiToConfig();
            ConfigHelper.Save(_appConfig);
        }

        private AppConfig _appConfig = new AppConfig();

        /// <summary>将配置应用到所有 UI 控件</summary>
        private void LoadConfigToUi()
        {
            _appConfig = ConfigHelper.Load();

            // ===== Tab 1：不恢复文件路径，只恢复筛选和清理状态 =====
            cbxFilterColumn.Text = _appConfig.FilterColumn;
            txtFilterValue.Text = _appConfig.FilterValue;

            lstCleanConditions.Items.Clear();
            foreach (var cond in _appConfig.CleanConditions)
                lstCleanConditions.Items.Add(cond);
            cbxRemoveDuplicateSN.Checked = _appConfig.RemoveDuplicateSN;
            cbxKeepOnlyCommonCols.Checked = _appConfig.KeepOnlyCommonCols;

            // ===== Tab 2：不恢复文件路径，只恢复选项状态 =====
            rdoFromExcel.Checked = _appConfig.SnFromExcel;
            rdoFromTxt.Checked = !_appConfig.SnFromExcel;
            rdoAllDrives.Checked = _appConfig.SearchAllDrives;
            rdoSpecifiedDir.Checked = !_appConfig.SearchAllDrives;

            // ===== Tab 3：原有设置 =====
            chkShowExt.Checked = _appConfig.ShowExt;
            chkShowHidden.Checked = _appConfig.ShowHidden;
            chkHibernateOff.Checked = _appConfig.HibernateOff;
            chkCleanTemp.Checked = _appConfig.CleanTemp;
            chkDesktopIcons.Checked = _appConfig.DesktopIcons;

            // ===== Tab 3：PPTX 新增设置 =====
            chkUsbPowerDisabled.Checked = _appConfig.UsbPowerDisabled;
            chkPciePowerDisabled.Checked = _appConfig.PciePowerDisabled;
            chkNetworkEbusOnly.Checked = _appConfig.NetworkEbusOnly;
            chkReceiveBuffersMax.Checked = _appConfig.ReceiveBuffersMax;
            chkInterruptModerationExtreme.Checked = _appConfig.InterruptModerationExtreme;
            chkBiosPcieAutoReminder.Checked = _appConfig.BiosPcieAutoReminder;

            // ===== Tab 3：新电脑必做设置 =====
            chkWinUpdate.Checked = _appConfig.WinUpdate;
            chkRestorePoint.Checked = _appConfig.RestorePoint;
            chkDefender.Checked = _appConfig.Defender;
            chkRemoveBloatware.Checked = _appConfig.RemoveBloatware;
            chkPowerHighPerf.Checked = _appConfig.PowerHighPerf;
            chkUAC.Checked = _appConfig.UAC;
            chkRecoveryDrive.Checked = _appConfig.RecoveryDrive;

            // ===== Tab 3：系统优化 =====
            chkOptSysMain.Checked = _appConfig.OptSysMain;
            chkOptWindowsSearch.Checked = _appConfig.OptWindowsSearch;
            chkOptTelemetry.Checked = _appConfig.OptTelemetry;
            chkOptBackgroundApps.Checked = _appConfig.OptBackgroundApps;
            chkOptOneDrive.Checked = _appConfig.OptOneDrive;
            chkOptVisualEffects.Checked = _appConfig.OptVisualEffects;
            chkOptTips.Checked = _appConfig.OptTips;
            chkOptLockScreen.Checked = _appConfig.OptLockScreen;

            // ===== Tab 3：隐私选项 =====
            chkPrivAdId.Checked = _appConfig.PrivAdId;
            chkPrivLocation.Checked = _appConfig.PrivLocation;
            chkPrivDiagnostics.Checked = _appConfig.PrivDiagnostics;
            chkPrivSpeech.Checked = _appConfig.PrivSpeech;

            // ===== Tab 3：快速启动 =====
            chkFastStartup.Checked = _appConfig.FastStartup;
        }

        /// <summary>从所有 UI 控件读取值到配置对象（不保存文件路径）</summary>
        private void SaveUiToConfig()
        {
            _appConfig = new AppConfig();

            // ===== Tab 1：只保存筛选和清理状态，不保存路径 =====
            _appConfig.FilterColumn = cbxFilterColumn.Text.Trim();
            _appConfig.FilterValue = txtFilterValue.Text.Trim();

            _appConfig.CleanConditions = new List<string>();
            foreach (var item in lstCleanConditions.Items)
                _appConfig.CleanConditions.Add(item.ToString());
            _appConfig.RemoveDuplicateSN = cbxRemoveDuplicateSN.Checked;
            _appConfig.KeepOnlyCommonCols = cbxKeepOnlyCommonCols.Checked;

            // ===== Tab 2：只保存选项状态，不保存路径 =====
            _appConfig.SnFromExcel = rdoFromExcel.Checked;
            _appConfig.SearchAllDrives = rdoAllDrives.Checked;

            // ===== Tab 3：原有设置 =====
            _appConfig.ShowExt = chkShowExt.Checked;
            _appConfig.ShowHidden = chkShowHidden.Checked;
            _appConfig.HibernateOff = chkHibernateOff.Checked;
            _appConfig.CleanTemp = chkCleanTemp.Checked;
            _appConfig.DesktopIcons = chkDesktopIcons.Checked;

            // ===== Tab 3：PPTX 新增设置 =====
            _appConfig.UsbPowerDisabled = chkUsbPowerDisabled.Checked;
            _appConfig.PciePowerDisabled = chkPciePowerDisabled.Checked;
            _appConfig.NetworkEbusOnly = chkNetworkEbusOnly.Checked;
            _appConfig.ReceiveBuffersMax = chkReceiveBuffersMax.Checked;
            _appConfig.InterruptModerationExtreme = chkInterruptModerationExtreme.Checked;
            _appConfig.BiosPcieAutoReminder = chkBiosPcieAutoReminder.Checked;

            // ===== Tab 3：新电脑必做设置 =====
            _appConfig.WinUpdate = chkWinUpdate.Checked;
            _appConfig.RestorePoint = chkRestorePoint.Checked;
            _appConfig.Defender = chkDefender.Checked;
            _appConfig.RemoveBloatware = chkRemoveBloatware.Checked;
            _appConfig.PowerHighPerf = chkPowerHighPerf.Checked;
            _appConfig.UAC = chkUAC.Checked;
            _appConfig.RecoveryDrive = chkRecoveryDrive.Checked;

            // ===== Tab 3：系统优化 =====
            _appConfig.OptSysMain = chkOptSysMain.Checked;
            _appConfig.OptWindowsSearch = chkOptWindowsSearch.Checked;
            _appConfig.OptTelemetry = chkOptTelemetry.Checked;
            _appConfig.OptBackgroundApps = chkOptBackgroundApps.Checked;
            _appConfig.OptOneDrive = chkOptOneDrive.Checked;
            _appConfig.OptVisualEffects = chkOptVisualEffects.Checked;
            _appConfig.OptTips = chkOptTips.Checked;
            _appConfig.OptLockScreen = chkOptLockScreen.Checked;

            // ===== Tab 3：隐私选项 =====
            _appConfig.PrivAdId = chkPrivAdId.Checked;
            _appConfig.PrivLocation = chkPrivLocation.Checked;
            _appConfig.PrivDiagnostics = chkPrivDiagnostics.Checked;
            _appConfig.PrivSpeech = chkPrivSpeech.Checked;

            // ===== Tab 3：快速启动 =====
            _appConfig.FastStartup = chkFastStartup.Checked;
        }

        /// <summary>
        /// 只保留所有行都有数据的列，去掉全空列
        /// </summary>
        private DataTable KeepOnlyCommonColumns(DataTable table)
        {
            if (table.Rows.Count == 0) return table;

            var colsToRemove = new List<DataColumn>();

            foreach (DataColumn col in table.Columns)
            {
                bool allEmpty = true;
                foreach (DataRow row in table.Rows)
                {
                    if (row.RowState == DataRowState.Deleted) continue;
                    object val = row[col];
                    if (val != null && val != DBNull.Value && !string.IsNullOrWhiteSpace(val.ToString()))
                    {
                        allEmpty = false;
                        break;
                    }
                }
                if (allEmpty) colsToRemove.Add(col);
            }

            foreach (var col in colsToRemove)
                table.Columns.Remove(col);

            return table;
        }
    }
}
