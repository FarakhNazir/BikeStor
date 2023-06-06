namespace BikeStor.Controllers
{
    internal class ConsolidatedChild
    {
        public ConsolidatedChild()
        {
        }

        public int ProductId { get; set; }
        public object Count_of_Products { get; set; }
        public decimal Unit_Price { get; set; }
        public decimal Total_price { get; set; }
    }
}