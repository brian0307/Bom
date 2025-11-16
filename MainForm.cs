using System;
using System.Collections.Generic;
using System.Data;
using System.Linq;
using System.Windows.Forms;
using ClosedXML.Excel;

namespace BomDfsApp
{
    public class MainForm : Form
    {
        private Button btnLoadBom;
        private Button btnExpandAll;
        private Button btnCollapseAll;
        private Button btnExport;
        private DataGridView dgv;
        private Label lblInfo;

        // 特性編碼顯示
        private Label lblFeature;
        private TextBox txtFeature;

        // BOM 資料
        private DataTable _bomRaw = new();      // 讀檔＋過濾失效日期後
        private DataTable _bomSorted = new();   // DFS 排序後 (完整炸開)
        private Dictionary<string, List<DataRow>> _childrenByParent = new();

        // 欄位名稱
        private readonly string COL_PARENT = "PARENT";
        private readonly string COL_CHILD = "CHILD";
        private readonly string COL_SEQ = "組合項次";
        private readonly string COL_FULLPATH = "FULL_PATH";
        private readonly string COL_EXPIRE = "失效日期";
        private readonly string COL_FEATURE = "特性編碼";
        private readonly string COL_LEVEL = "階層";

        // 特性編碼值
        private string _featureCode = "";

        // DataGridView 第一欄欄名 (展開 / 收合)
        private readonly string COL_EXPAND = "_EXPAND_";

        // 用來記錄每一列對應的 DataRow & 是否已展開
        private class BomNodeState
        {
            public DataRow Row { get; set; } = null!;
            public bool IsExpanded { get; set; } = false;
        }

        public MainForm()
        {
            Text = "BOM 樹狀展開工具";
            Width = 1200;
            Height = 720;
            StartPosition = FormStartPosition.CenterScreen;

            btnLoadBom = new Button
            {
                Text = "載入 BOM",
                Left = 20,
                Top = 20,
                Width = 120,
                Height = 30
            };
            btnExpandAll = new Button
            {
                Text = "全部展開",
                Left = 160,
                Top = 20,
                Width = 120,
                Height = 30
            };
            btnCollapseAll = new Button
            {
                Text = "全部收合",
                Left = 300,
                Top = 20,
                Width = 120,
                Height = 30
            };
            btnExport = new Button
            {
                Text = "匯出 Excel",
                Left = 440,
                Top = 20,
                Width = 120,
                Height = 30
            };

            lblFeature = new Label
            {
                Left = 580,
                Top = 24,
                Width = 80,
                Height = 20,
                Text = "特性編碼:"
            };
            txtFeature = new TextBox
            {
                Left = 660,
                Top = 20,
                Width = 200,
                Height = 24,
                ReadOnly = true
            };

            lblInfo = new Label
            {
                Left = 20,
                Top = 60,
                Width = 1100,
                Height = 30,
                Text = "請先載入 BOM（階層 / PARENT / CHILD / 組合項次 / FULL_PATH / 失效日期 / 特性編碼 ...）。"
            };

            dgv = new DataGridView
            {
                Left = 20,
                Top = 100,
                Width = 1140,
                Height = 560,
                ReadOnly = true,
                AllowUserToAddRows = false,
                AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.DisplayedCells
            };

            Controls.Add(btnLoadBom);
            Controls.Add(btnExpandAll);
            Controls.Add(btnCollapseAll);
            Controls.Add(btnExport);
            Controls.Add(lblFeature);
            Controls.Add(txtFeature);
            Controls.Add(lblInfo);
            Controls.Add(dgv);

            btnLoadBom.Click += BtnLoadBom_Click;
            btnExpandAll.Click += BtnExpandAll_Click;
            btnCollapseAll.Click += BtnCollapseAll_Click;
            btnExport.Click += BtnExport_Click;
            dgv.CellClick += Dgv_CellClick;
        }

        // ================== 1. 載入 BOM ==================
        private void BtnLoadBom_Click(object? sender, EventArgs e)
        {
            using var ofd = new OpenFileDialog
            {
                Filter = "Excel Files|*.xlsx;*.xlsm;*.xls",
                Title = "選擇 Bom.xlsx"
            };
            if (ofd.ShowDialog() != DialogResult.OK) return;

            try
            {
                _bomRaw = LoadBomExcel(ofd.FileName);

                // 先用 DFS 炸 BOM，得到完整排序
                var order = BuildDfsOrder(_bomRaw);
                _bomSorted = ApplyOrder(_bomRaw, order);

                // 建立 Parent -> Children 對照
                BuildChildrenMap();

                // 建立欄位 (手動建立，方便做樹狀＋隱藏 helper 欄位)
                BuildGridColumns();

                // 一開始只顯示 Level = 1
                ShowOnlyLevel1();

                txtFeature.Text = _featureCode;

                lblInfo.Text = $"已載入：{_bomSorted.Rows.Count} 列 (已排除失效日期有值的列)。目前只顯示階層 = 1。";
            }
            catch (Exception ex)
            {
                MessageBox.Show("載入 / 處理 BOM 失敗： " + ex.Message);
            }
        }

        // ================== 2. 全部展開 ==================
        private void BtnExpandAll_Click(object? sender, EventArgs e)
        {
            if (_bomSorted == null || _bomSorted.Rows.Count == 0)
            {
                MessageBox.Show("請先載入 BOM。");
                return;
            }

            dgv.Rows.Clear();

            foreach (DataRow row in _bomSorted.Rows)
            {
                AddGridRowFromDataRow(row);
            }

            // 所有有子階的節點設成「已展開」，符號顯示為 -
            for (int i = 0; i < dgv.Rows.Count; i++)
            {
                var gr = dgv.Rows[i];
                if (gr.Tag is BomNodeState state)
                {
                    string childKey = state.Row[COL_CHILD]?.ToString()?.Trim() ?? "";
                    if (!string.IsNullOrEmpty(childKey) &&
                        _childrenByParent.ContainsKey(childKey) &&
                        _childrenByParent[childKey].Count > 0)
                    {
                        state.IsExpanded = true;
                        gr.Cells[COL_EXPAND].Value = "-";
                    }
                }
            }

            lblInfo.Text = $"全部展開：顯示 {_bomSorted.Rows.Count} 列。";
        }

        // ================== 3. 全部收合（只剩階層 = 1） ==================
        private void BtnCollapseAll_Click(object? sender, EventArgs e)
        {
            if (_bomSorted == null || _bomSorted.Rows.Count == 0)
            {
                MessageBox.Show("請先載入 BOM。");
                return;
            }

            ShowOnlyLevel1();
            lblInfo.Text = "已收合：只顯示階層 = 1。";
        }

        // ================== 4. 匯出 Excel（同之前邏輯） ==================
        private void BtnExport_Click(object? sender, EventArgs e)
        {
            if (_bomSorted == null || _bomSorted.Rows.Count == 0)
            {
                MessageBox.Show("沒有資料可以匯出，請先載入 BOM。");
                return;
            }

            using var sfd = new SaveFileDialog
            {
                Filter = "Excel Files|*.xlsx",
                FileName = "Bom_Dfs_Result.xlsx"
            };
            if (sfd.ShowDialog() != DialogResult.OK) return;

            try
            {
                ExportToExcelWithoutHelperCols(_bomSorted, sfd.FileName);
                MessageBox.Show("匯出完成！");
            }
            catch (Exception ex)
            {
                MessageBox.Show("匯出失敗： " + ex.Message);
            }
        }

        // ================== 5. Grid 點擊 + / - 展開 / 收合 ==================
        private void Dgv_CellClick(object? sender, DataGridViewCellEventArgs e)
        {
            if (e.RowIndex < 0 || e.ColumnIndex < 0) return;

            var col = dgv.Columns[e.ColumnIndex];
            if (col.Name != COL_EXPAND) return;      // 只處理第一欄 (+/-)

            var row = dgv.Rows[e.RowIndex];
            if (row.Tag is not BomNodeState state) return;

            string childKey = state.Row[COL_CHILD]?.ToString()?.Trim() ?? "";
            if (string.IsNullOrEmpty(childKey) ||
                !_childrenByParent.ContainsKey(childKey) ||
                _childrenByParent[childKey].Count == 0)
            {
                return; // 沒有子階不用動
            }

            if (state.IsExpanded)
                CollapseRow(e.RowIndex);
            else
                ExpandRow(e.RowIndex);
        }

        // 展開某一列 (只展開下一階)
        private void ExpandRow(int rowIndex)
        {
            var row = dgv.Rows[rowIndex];
            if (row.Tag is not BomNodeState state) return;

            string childKey = state.Row[COL_CHILD]?.ToString()?.Trim() ?? "";
            if (!_childrenByParent.TryGetValue(childKey, out var children) || children.Count == 0)
                return;

            int insertIndex = rowIndex + 1;
            foreach (var childRow in children)
            {
                // 如果已經在畫面上，就不要重複插入 (簡單判斷：看 Tag.Row 是否已存在)
                bool alreadyVisible = false;
                foreach (DataGridViewRow gr in dgv.Rows)
                {
                    if (gr.Tag is BomNodeState s && s.Row == childRow)
                    {
                        alreadyVisible = true;
                        break;
                    }
                }
                if (alreadyVisible) continue;

                AddGridRowFromDataRow(childRow, insertIndex);
                insertIndex++;
            }

            state.IsExpanded = true;
            row.Cells[COL_EXPAND].Value = "-";
        }

        // 收合某一列 (移除所有後續、階層比自己大的列)
        private void CollapseRow(int rowIndex)
        {
            var row = dgv.Rows[rowIndex];
            if (row.Tag is not BomNodeState state) return;

            int myLevel = ParseInt(state.Row[COL_LEVEL]);
            int i = rowIndex + 1;

            while (i < dgv.Rows.Count)
            {
                var gr = dgv.Rows[i];
                if (gr.Tag is not BomNodeState s) { i++; continue; }

                int level = ParseInt(s.Row[COL_LEVEL]);
                if (level <= myLevel)
                    break;      // 到同層或更上層就停

                dgv.Rows.RemoveAt(i); // 不用 i++，因為移除後下一列會補上來
            }

            state.IsExpanded = false;
            // 收合後，恢復成有子階就顯示 '+'
            string childKey = state.Row[COL_CHILD]?.ToString()?.Trim() ?? "";
            if (!string.IsNullOrEmpty(childKey) &&
                _childrenByParent.ContainsKey(childKey) &&
                _childrenByParent[childKey].Count > 0)
            {
                row.Cells[COL_EXPAND].Value = "+";
            }
            else
            {
                row.Cells[COL_EXPAND].Value = "";
            }
        }

        // ================== 內部小工具 ==================

        // 讀 BOM Excel：過濾失效日期、抓特性編碼
        private DataTable LoadBomExcel(string path)
        {
            using var wb = new XLWorkbook(path);
            var ws = wb.Worksheet(1);  // 第一個工作表

            var dt = new DataTable();
            bool firstRow = true;

            foreach (var row in ws.RowsUsed())
            {
                if (firstRow)
                {
                    foreach (var cell in row.CellsUsed())
                        dt.Columns.Add(cell.GetString().Trim());
                    firstRow = false;
                }
                else
                {
                    var newRow = dt.NewRow();
                    int colIndex = 0;
                    foreach (var cell in row.Cells(1, dt.Columns.Count))
                    {
                        newRow[colIndex++] = cell.Value;
                    }
                    dt.Rows.Add(newRow);
                }
            }

            if (!dt.Columns.Contains(COL_PARENT) ||
                !dt.Columns.Contains(COL_CHILD) ||
                !dt.Columns.Contains(COL_SEQ) ||
                !dt.Columns.Contains(COL_LEVEL))
                throw new Exception("BOM 檔缺少必要欄位：階層 / PARENT / CHILD / 組合項次。");

            if (!dt.Columns.Contains(COL_FULLPATH))
                dt.Columns.Add(COL_FULLPATH, typeof(string));

            bool hasExpire = dt.Columns.Contains(COL_EXPIRE);
            bool hasFeature = dt.Columns.Contains(COL_FEATURE);

            // 抓特性編碼
            _featureCode = "";
            if (hasFeature)
            {
                foreach (DataRow r in dt.Rows)
                {
                    var v = r[COL_FEATURE]?.ToString()?.Trim();
                    if (!string.IsNullOrEmpty(v))
                    {
                        _featureCode = v;
                        break;
                    }
                }
            }

            // 失效日期有值的列全部刪掉
            if (hasExpire)
            {
                for (int i = dt.Rows.Count - 1; i >= 0; i--)
                {
                    var v = dt.Rows[i][COL_EXPIRE];
                    if (v != null && v != DBNull.Value &&
                        !string.IsNullOrWhiteSpace(v.ToString()))
                    {
                        dt.Rows.RemoveAt(i);
                    }
                }
            }

            dt.AcceptChanges();
            return dt;
        }

        // 建 DFS 順序 (跟之前版本一樣)
        private List<int> BuildDfsOrder(DataTable dt)
        {
            var edges = new List<BomEdge>();
            for (int i = 0; i < dt.Rows.Count; i++)
            {
                var row = dt.Rows[i];
                string parent = row[COL_PARENT]?.ToString()?.Trim() ?? "";
                string child = row[COL_CHILD]?.ToString()?.Trim() ?? "";
                if (string.IsNullOrEmpty(parent) || string.IsNullOrEmpty(child))
                    continue;

                int seq = int.MaxValue;
                var seqObj = row[COL_SEQ];
                if (seqObj != null && seqObj != DBNull.Value &&
                    int.TryParse(seqObj.ToString().Trim(), out var tmp))
                    seq = tmp;

                edges.Add(new BomEdge
                {
                    Parent = parent,
                    Child = child,
                    Seq = seq,
                    RowIndex = i
                });
            }

            var map = edges
                .GroupBy(e => e.Parent)
                .ToDictionary(
                    g => g.Key,
                    g => g.OrderBy(x => x.Seq).ThenBy(x => x.Child).ToList()
                );

            var parents = new HashSet<string>(edges.Select(e => e.Parent));
            var childs = new HashSet<string>(edges.Select(e => e.Child));
            var roots = parents.Except(childs).ToList();

            if (!roots.Any())
                throw new Exception("找不到 Root（成品），請確認 BOM 資料。");

            var order = new List<int>();
            var visited = new HashSet<string>();

            foreach (var root in roots.OrderBy(r => r))
            {
                DfsNode(root, map, visited, order);
            }

            return order;
        }

        private void DfsNode(
            string parent,
            Dictionary<string, List<BomEdge>> map,
            HashSet<string> visitedInPath,
            List<int> order)
        {
            if (visitedInPath.Contains(parent))
                return;

            visitedInPath.Add(parent);

            if (map.TryGetValue(parent, out var children))
            {
                foreach (var edge in children)
                {
                    order.Add(edge.RowIndex);
                    DfsNode(edge.Child, map, visitedInPath, order);
                }
            }

            visitedInPath.Remove(parent);
        }

        // 依 DFS 順序重建 DataTable
        private DataTable ApplyOrder(DataTable dt, List<int> order)
        {
            var result = dt.Clone();

            foreach (int idx in order)
            {
                if (idx >= 0 && idx < dt.Rows.Count)
                    result.ImportRow(dt.Rows[idx]);
            }

            return result;
        }

        // 建 Parent -> Children 對照
        private void BuildChildrenMap()
        {
            _childrenByParent = new Dictionary<string, List<DataRow>>(StringComparer.OrdinalIgnoreCase);

            foreach (DataRow row in _bomSorted.Rows)
            {
                string parent = row[COL_PARENT]?.ToString()?.Trim() ?? "";
                string child = row[COL_CHILD]?.ToString()?.Trim() ?? "";
                if (string.IsNullOrEmpty(parent) || string.IsNullOrEmpty(child))
                    continue;

                if (!_childrenByParent.TryGetValue(parent, out var list))
                {
                    list = new List<DataRow>();
                    _childrenByParent[parent] = list;
                }

                list.Add(row);   // 依 _bomSorted 的順序加入
            }
        }

        // 建立 DataGridView 欄位
        private void BuildGridColumns()
        {
            dgv.Columns.Clear();

            // 第一欄：展開 / 收合用
            var expandCol = new DataGridViewTextBoxColumn
            {
                Name = COL_EXPAND,
                HeaderText = "",
                Width = 30,
                ReadOnly = true
            };
            dgv.Columns.Add(expandCol);

            // 其他欄位：依 _bomSorted.Columns 建立，但排除 helper 欄位
            foreach (DataColumn col in _bomSorted.Columns)
            {
                if (string.Equals(col.ColumnName, COL_PARENT, StringComparison.OrdinalIgnoreCase) ||
                    string.Equals(col.ColumnName, COL_SEQ, StringComparison.OrdinalIgnoreCase) ||
                    string.Equals(col.ColumnName, COL_FULLPATH, StringComparison.OrdinalIgnoreCase) ||
                    string.Equals(col.ColumnName, COL_FEATURE, StringComparison.OrdinalIgnoreCase))
                {
                    continue;   // 不顯示的欄位
                }

                var gridCol = new DataGridViewTextBoxColumn
                {
                    Name = col.ColumnName,
                    HeaderText = col.ColumnName,
                    ReadOnly = true
                };

                // CHILD 欄位改成「料號」
                if (string.Equals(col.ColumnName, COL_CHILD, StringComparison.OrdinalIgnoreCase))
                    gridCol.HeaderText = "料號";

                dgv.Columns.Add(gridCol);
            }
        }

        // 新增一列到 DataGridView（附帶 Tag / 展開符號）
        private void AddGridRowFromDataRow(DataRow srcRow, int? insertIndex = null)
        {
            int rowIndex;
            if (insertIndex.HasValue)
            {
                // Insert 不會回傳 int，所以先插入，再把 rowIndex 設成插入點
                dgv.Rows.Insert(insertIndex.Value, 1);
                rowIndex = insertIndex.Value;
            }
            else
            {
                rowIndex = dgv.Rows.Add();
            }

            var gr = dgv.Rows[rowIndex];

            var state = new BomNodeState { Row = srcRow, IsExpanded = false };
            gr.Tag = state;

            string childKey = srcRow[COL_CHILD]?.ToString()?.Trim() ?? "";
            bool hasChildren = !string.IsNullOrEmpty(childKey) &&
                               _childrenByParent.ContainsKey(childKey) &&
                               _childrenByParent[childKey].Count > 0;

            gr.Cells[COL_EXPAND].Value = hasChildren ? "+" : "";

            // 填其他欄位
            foreach (DataGridViewColumn col in dgv.Columns)
            {
                if (col.Name == COL_EXPAND) continue;

                var value = srcRow[col.Name];
                gr.Cells[col.Name].Value =
                    value == null || value == DBNull.Value ? "" : value.ToString();
            }
        }

        // 只顯示階層 = 1
        private void ShowOnlyLevel1()
        {
            dgv.Rows.Clear();

            foreach (DataRow row in _bomSorted.Rows)
            {
                int level = ParseInt(row[COL_LEVEL]);
                if (level == 1)
                {
                    AddGridRowFromDataRow(row);
                }
            }
        }

        private int ParseInt(object? v)
        {
            if (v == null || v == DBNull.Value) return int.MaxValue;
            var s = v.ToString()?.Trim();
            if (int.TryParse(s, out var n)) return n;
            return int.MaxValue;
        }

        // 匯出（排除 PARENT / 組合項次 / FULL_PATH / 特性編碼）
        private void ExportToExcelWithoutHelperCols(DataTable dt, string path)
        {
            using var wb = new XLWorkbook();
            var ws = wb.AddWorksheet("BOM_DFS");

            var exportColumns = dt.Columns
                .Cast<DataColumn>()
                .Where(c =>
                    !string.Equals(c.ColumnName, COL_PARENT, StringComparison.OrdinalIgnoreCase) &&
                    !string.Equals(c.ColumnName, COL_SEQ, StringComparison.OrdinalIgnoreCase) &&
                    !string.Equals(c.ColumnName, COL_FULLPATH, StringComparison.OrdinalIgnoreCase) &&
                    !string.Equals(c.ColumnName, COL_FEATURE, StringComparison.OrdinalIgnoreCase))
                .ToList();

            int col = 1;
            foreach (var c in exportColumns)
            {
                ws.Cell(1, col).Value = c.ColumnName;
                col++;
            }

            int rowIndex = 2;
            foreach (DataRow row in dt.Rows)
            {
                col = 1;
                foreach (var c in exportColumns)
                {
                    var v = row[c.ColumnName];
                    ws.Cell(rowIndex, col).Value =
                        v == null || v == DBNull.Value ? "" : v.ToString();
                    col++;
                }
                rowIndex++;
            }

            ws.Columns().AdjustToContents();
            wb.SaveAs(path);
        }
    }
}
