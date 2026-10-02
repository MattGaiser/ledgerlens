namespace LedgerLens.Core
{
    public static class ProductInfo
    {
        public static string Version => typeof(ProductInfo).Assembly.GetName().Version!.ToString(3);
    }
}
