# Simple e-commerce SKU calculator

product_name = "Wireless Headphones"
selling_price = 50.00
cost_per_unit = 30.00
units_sold = 1000
website_visitors = 25000

revenue = selling_price * units_sold
total_cost = cost_per_unit * units_sold
gross_profit = revenue - total_cost
profit_margin = (gross_profit / revenue) * 100
conversion_rate = (units_sold / website_visitors) * 100

print(f"Product: {product_name}")
print(f"Units sold: {units_sold:,}")
print(f"Sales revenue: ${revenue:,.2f}")
print(f"Gross profit: ${gross_profit:,.2f}")
print(f"Profit margin: {profit_margin:.2f}%")
print(f"Conversion rate: {conversion_rate:.2f}%")
# Simple e-commerce SKU calculator

product_name = "Wireless Headphones"
selling_price = 50.00
cost_per_unit = 30.00
units_sold = 1000
website_visitors = 25000
target_net_profit_margin = 20 / 100

revenue = selling_price * units_sold
total_product_cost = cost_per_unit * units_sold
gross_profit = revenue - total_product_cost
gross_profit_margin = (gross_profit / revenue) * 100
conversion_rate = (units_sold / website_visitors) * 100

# Net profit at a 20% target margin
net_profit = revenue * target_net_profit_margin
maximum_total_expenses = revenue - net_profit

print(f"Product: {product_name}")
print(f"Units sold: {units_sold:,}")
print(f"Sales revenue: ${revenue:,.2f}")
print(f"Gross profit: ${gross_profit:,.2f}")
print(f"Gross profit margin: {gross_profit_margin:.2f}%")
print(f"Target net profit margin: {target_net_profit_margin:.0%}")
print(f"Net profit: ${net_profit:,.2f}")
print(f"Maximum total expenses: ${maximum_total_expenses:,.2f}")
print(f"Conversion rate: {conversion_rate:.2f}%")