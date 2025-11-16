namespace BomDfsApp
{
    // 用來建樹 / DFS 的 edge
    public class BomEdge
    {
        public string Parent { get; set; } = "";
        public string Child { get; set; } = "";
        public int Seq { get; set; }   // 組合項次
        public int RowIndex { get; set; }   // 在 DataTable 裡的列索引
    }
}
