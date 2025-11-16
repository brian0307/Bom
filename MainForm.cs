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
        private Button btnExplode;
        private Button btnExport;
        private DataGridView dgv;
        private Label lblInfo;

        // 特性編碼顯示用
        private Label lblFeature;
        private TextBox txtFeature;

        private DataTable _rawBom = new();     // 原始(已過濾失效) BOM
        private DataTable _sortedBom = new();  // DFS 排序後

        // 欄位名稱常數
        private readonly string COL_PARENT = "PARENT";
        private readonly string COL_CHILD = "CHILD";
        private readonly string COL_SEQ = "組合項次";
        private readonly string COL_FULLPATH = "FULL_PATH";
        private readonly string COL_EXPIRE = "失效日期";
        private readonly string COL_FEATURE = "特性編碼";

        // 特性編碼值（從 BOM 中抓第一個）
        private string _featureCode = "";

        public MainForm()
        {
            Text = "BOM DFS 展開工具";
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
            btnExplode = new Button
            {
                Text = "炸 BOM (DFS)",
                Left = 160,
                Top = 20,
                Width = 140,
                Height = 30
            };
            btnExport = new Button
            {
                Text = "匯出 Excel",
                Left = 320,
                Top = 20,
                Width = 140,
                Height = 30
            };

            lblFeature = new Label
            {
                Left = 480,
                Top = 24,
                Width = 80,
                Height = 20,
                Text = "特性編碼:"
            };
            txtFeature = new TextBox
            {
                Left = 560,
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
                Text = "請先載入 Bom.xlsx（PARENT / CHILD / 組合項次 / FULL_PATH / IS_LEAF / 失效日期 / 特性編碼）。"
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
            Controls.Add(btnExplode);
            Controls.Add(btnExport);
            Controls.Add(lblFeature);
            Controls.Add(txtFeature);
            Controls.Add(lblInfo);
            Controls.Add(dgv);

            btnLoadBom.Click += BtnLoadBom_Click;
            btnExplode.Click += BtnExplode_Click;
            btnExport.Click += BtnExport_Click;
        }

        // 1. 載入 BOM Excel
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
                _rawBom = LoadBomExcel(ofd.FileName);
                _sortedBom = _rawBom.Copy(); // 先顯示未排序
                dgv.DataSource = _sortedBom;

                HideHelperColumnsInGrid();

                txtFeature.Text = _featureCode; // 顯示特性編碼

                lblInfo.Text = $"已載入：{_rawBom.Rows.Count} 筆（已排除失效日期有值的列）。按「炸 BOM」會依 Parent/Child + 組合項次做 DFS 排序。";
            }
            catch (Exception ex)
            {
                MessageBox.Show("載入 BOM 失敗： " + ex.Message);
            }
        }

        // 2. 炸 BOM：依 Parent/Child + 組合項次 DFS
        private void BtnExplode_Click(object? sender, EventArgs e)
        {
            if (_rawBom == null || _rawBom.Rows.Count == 0)
            {
                MessageBox.Show("請先載入 BOM。");
                return;
            }

            try
            {
                var order = BuildDfsOrder(_rawBom);
                _sortedBom = ApplyOrder(_rawBom, order);

                dgv.DataSource = _sortedBom;
                HideHelperColumnsInGrid();

                lblInfo.Text = $"炸 BOM 完成：共 {_sortedBom.Rows.Count} 列，已依 DFS + 組合項次排序。";
            }
            catch (Exception ex)
            {
                MessageBox.Show("炸 BOM 時發生錯誤： " + ex.Message);
            }
        }

        // 3. 匯出結果 Excel（不含 PARENT / 組合項次 / FULL_PATH / IS_LEAF / 特性編碼）
        private void BtnExport_Click(object? sender, EventArgs e)
        {
            if (_sortedBom == null || _sortedBom.Rows.Count == 0)
            {
                MessageBox.Show("沒有資料可匯出，請先炸 BOM。");
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
                ExportToExcelWithoutHelperCols(_sortedBom, sfd.FileName);
                MessageBox.Show("匯出完成！");
            }
            catch (Exception ex)
            {
                MessageBox.Show("匯出失敗： " + ex.Message);
            }
        }

        // ========= 讀 Excel =========
        private DataTable LoadBomExcel(string path)
        {
            using var wb = new XLWorkbook(path);
            var ws = wb.Worksheet(1);  // 預設第一個工作表「工作表1」

            var dt = new DataTable();
            bool firstRow = true;

            foreach (var row in ws.RowsUsed())
            {
                if (firstRow)
                {
                    // 表頭
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

            // 檢查必備欄位
            if (!dt.Columns.Contains(COL_PARENT) ||
                !dt.Columns.Contains(COL_CHILD) ||
                !dt.Columns.Contains(COL_SEQ))
                throw new Exception("BOM 檔缺少必要欄位：PARENT / CHILD / 組合項次。");

            // FULL_PATH / IS_LEAF 如果沒有，就補空欄位
            if (!dt.Columns.Contains(COL_FULLPATH))
                dt.Columns.Add(COL_FULLPATH, typeof(string));

            bool hasExpire = dt.Columns.Contains(COL_EXPIRE);
            bool hasFeature = dt.Columns.Contains(COL_FEATURE);

            // 先抓特性編碼（找第一個非空）
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

            // 若有 失效日期 欄位，將「失效日期不是空」的列剔除
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

        // ========= 建立 DFS 順序 =========
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

            // Parent -> List<Edge>（同一 parent 下按 組合項次 排序）
            var map = edges
                .GroupBy(e => e.Parent)
                .ToDictionary(
                    g => g.Key,
                    g => g.OrderBy(x => x.Seq).ThenBy(x => x.Child).ToList()
                );

            // 找 root：出現在 PARENT 但沒出現在 CHILD 的節點
            var parents = new HashSet<string>(edges.Select(e => e.Parent));
            var childs = new HashSet<string>(edges.Select(e => e.Child));
            var roots = parents.Except(childs).ToList();

            if (!roots.Any())
                throw new Exception("找不到 Root（成品），請確認 BOM 資料。");

            var order = new List<int>();
            var visitedInPath = new HashSet<string>();

            foreach (var root in roots.OrderBy(r => r))
            {
                DfsNode(root, map, visitedInPath, order);
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
                    // 先記錄這個 edge 對應的列索引
                    order.Add(edge.RowIndex);

                    // 再往下炸 child
                    DfsNode(edge.Child, map, visitedInPath, order);
                }
            }

            visitedInPath.Remove(parent);
        }

        // ========= 套用 DFS 順序到 DataTable =========
        private DataTable ApplyOrder(DataTable dt, List<int> order)
        {
            var result = dt.Clone(); // 複製欄位結構

            foreach (int idx in order)
            {
                if (idx >= 0 && idx < dt.Rows.Count)
                {
                    result.ImportRow(dt.Rows[idx]);
                }
            }

            return result;
        }

        // ========= 匯出（不含 5 個 helper 欄） =========
        private void ExportToExcelWithoutHelperCols(DataTable dt, string path)
        {
            using var wb = new XLWorkbook();
            var ws = wb.AddWorksheet("BOM_DFS");

            // 要輸出的欄位（排除：PARENT / 組合項次 / FULL_PATH / IS_LEAF / 特性編碼）
            var exportColumns = dt.Columns
                .Cast<DataColumn>()
                .Where(c =>
                    !string.Equals(c.ColumnName, COL_PARENT, StringComparison.OrdinalIgnoreCase) &&
                    !string.Equals(c.ColumnName, COL_SEQ, StringComparison.OrdinalIgnoreCase) &&
                    !string.Equals(c.ColumnName, COL_FULLPATH, StringComparison.OrdinalIgnoreCase) &&
                    !string.Equals(c.ColumnName, COL_FEATURE, StringComparison.OrdinalIgnoreCase))
                .ToList();

            // 表頭
            int col = 1;
            foreach (var c in exportColumns)
            {
                ws.Cell(1, col).Value = c.ColumnName;
                col++;
            }

            // 資料
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

        // ========= 隱藏 helper 欄位（只影響畫面，不影響資料） =========
        private void HideHelperColumnsInGrid()
        {
            // 隱藏不需要顯示的欄位
            foreach (var name in new[] { COL_PARENT, COL_SEQ, COL_FULLPATH, COL_FEATURE })
            {
                if (dgv.Columns.Contains(name))
                    dgv.Columns[name].Visible = false;
            }

            // 把 CHILD 欄位名稱改成「料號」
            if (dgv.Columns.Contains(COL_CHILD))
                dgv.Columns[COL_CHILD].HeaderText = "料號";
        }
    }
}
