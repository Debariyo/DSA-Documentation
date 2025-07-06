# DSA Data Analysis Capstone Project 1
The journey of becoming a professional data analyst...

## Project Topic: Amazon Product Review Analysis 

### Project Overview: Analysis of Product and Customer Review Data for Strategic Insights

This project focuses on analyzing product and customer review data sourced from Amazon. The objective is to generate actionable insights that support product development, refine marketing strategies, and improve customer engagement for our clients.

This involves the collection, cleaning, and exploratory analysis of product performance metrics and textual customer reviews. Using tools like Excel, and visualization platforms such as PivotTable, I identify trends, pain points, and opportunities within customer feedback.

Key areas of analysis include:

* **Product Rating Trends:** Understanding average ratings across categories and time to assess product satisfaction.
* **Sentiment Analysis of Reviews:** Categorizing reviews as positive, negative, or neutral to determine customer perception.
* **Keyword Extraction & Topic Modeling:** Identifying frequently mentioned features, issues, or benefits in customer reviews.
* **Top/Low Performing Products:** Comparing products by combining rating and review volume to identify high-impact performers or underperformers.
* **Customer Segmentation :** Grouping customers based on their review behavior or product preferences.

#### Data Source
The primary source of data used here is Amazon case study.xlsx. It was provided by the company. 

#### 🔧 Tools Used
* **Microsoft Excel:** For data cleaning, manipulation, and analysis [Download Here](https://www.microsoft.com/en-us/download)
* **PivotTables & PivotCharts:** To summarize and explore product ratings, review counts, and customer sentiment trends
* **Excel Dashboard:** For visualizing key performance indicators (KPIs) like top-rated products, review volume, and product performance metrics
* **Conditional Formatting & Data Validation:** To enhance data readability and accuracy

#### Data Cleaning and Preparation
In the initial phase of the Data Cleaning and preparation, we pwerform the following actions:
1. Data loading and inspection
2. Handling missing variables
3. Data Cleaning and formatting

#### Exploratory Data Analysis
EDA involved the exploration of the Data to answer some questions about the Data such as:
1. What is the average discount percentage by product category? 
2. How many products are listed under each category? 
3. What is the total number of reviews per category?  
4. Which products have the highest average ratings?

#### 📊 Data Analysis
The analysis was performed entirely in Microsoft Excel using built-in functions, PivotTables, and dashboard components. Below are examples of techniques and formulas used:
**📌 Sample Excel Functions Used:**
* `=AVERAGEIF(range, criteria, average_range)` – to calculate average rating per product category
* `=COUNTIF(range, criteria)` – to count specific review types (e.g., 1-star, 5-star)
* `=IF(condition, value_if_true, value_if_false)` – for conditional flagging (e.g., low-rated products)

**📌 PivotTable & PivotChart Usage:**
* Created PivotTables to summarize:
  * Total number of reviews by product
  * Average rating by category
  * Monthly review trends
* Used PivotCharts to visualize:
  * Rating distribution
  * Top 5 and bottom 5 products
  * Review volume over time

**📌 Dashboard Components:**
* Slicers added for filtering by:
  * Product category
  * Rating level
  * Time period
* Conditional formatting to highlight:
  * Products with average rating < 3
  * Products with highest review volume

#### Result Findings

#### 1. Product Distribution by Category:
* **Electronics** has the highest number of listed products with 476, followed closely by **Home&Kitchen** with 448 products, and then **Computers&Accessories** with 375 products.
* Categories like **Car&Motorbike**, **Health&PersonalCare**, and **Toys&Games** have a very small number of products (1 each), suggesting they are either niche categories or less represented in this dataset. **MusicalInstruments** and **HomeImprovement** also have very few products (2 each).
* The **Grand Total** shows 1337 distinct products across all listed categories.

#### 2. Average Discount Percentage by Product Category:
* **Computers&Accessories** leads with the highest average discount percentage at approximately 54.02%, indicating that products in this category are frequently offered with substantial discounts.
* **HomeImprovement** and **Health&PersonalCare** also show high average discounts at 57.5% and 53% respectively, although these categories have fewer products overall.
* **Electronics** has an average discount of about 50.83%.
* **OfficeProducts** has a significantly lower average discount of approximately 12.35%, suggesting that these products are less often heavily discounted.
* **Toys&Games** shows a 0% average discount, which might indicate no discounts were applied to the single product in this category in the dataset.
* The overall average discount across all products is about 47.69%.

#### 3. Total Number of Reviews per Category:
* **Electronics** boasts the highest sum of rating counts, with over 15.7 million reviews, demonstrating extremely high customer engagement and volume of sales in this category.
* **Computers&Accessories** follows with over 7.7 million reviews.
* **Home&Kitchen** has a substantial number of reviews with nearly 3 million.
* Other categories like **OfficeProducts**, **MusicalInstruments**, **HomeImprovement**, **Car&Motorbike**, **Health&PersonalCare**, and **Toys&Games** have significantly fewer reviews, consistent with their lower product counts.
* The **Grand Total** of reviews stands at over 26.7 million, highlighting the vast scale of customer feedback.

#### 4. Average Actual Price vs. Discounted Price by Category:
* **Electronics** products have the highest average actual price (₹9880.13) and average discounted price (₹5965.89), reflecting the higher value of products in this category. The significant difference between actual and discounted prices also indicates large average discounts.
* **Computers&Accessories** has average actual prices around ₹1683.62 and discounted prices around ₹842.65, again showing substantial price reductions.
* **Car&Motorbike**, **Health&PersonalCare**, **Home&Kitchen**, and **MusicalInstruments** also show a notable difference between their actual and discounted prices, indicating common discounting practices.
* **OfficeProducts** has relatively low average actual (₹397.19) and discounted prices (₹301.58), aligning with its lower discount percentage.

#### 5. Distribution of Product Ratings & Relation to Discount Levels:
* The "Average of Rating" across all products is approximately 4.096, indicating generally positive customer satisfaction.
* Looking at the "Average of Rating" by discount percentage bands:
    * Products with **91-100% discounts** have the highest average rating of approximately 4.22, suggesting that heavily discounted products tend to maintain high satisfaction.
    * Products with **0-10% discounts** also show a high average rating of approximately 4.21.
    * The lowest average rating of about 3.95 is observed for products with **81-90% discounts**.
    * This might suggest that moderate discounts (e.g., 21-70%) maintain consistently good ratings (around 4.06 to 4.15), while very high discounts (91-100%) or very low/no discounts (0-10%) also correlate with high average ratings. More detailed analysis would be needed to confirm causality.

#### 6. Number of Unique Products per Price Range Bucket:
* The majority of products fall into the **>₹500** price range with 1245 products, indicating a concentration of products at higher price points.
* The **₹200-₹500** range has 183 products.
* The **<₹200** range has the fewest products with 37.

#### 7. Categories with Highest Discounts:
* **Computers&Accessories** has products with maximum discounts reaching 94%.
* **Electronics** (91%) and **Home&Kitchen** (90%) also feature products with very high individual discounts.
* This indicates that these categories are prime areas where consumers can find deeply discounted products.

Which products have the highest average ratings?
Based on the dataset analysis, the product with the highest average ratings are as follows:
1.	Syncwire LTG to USB Cable – Average Rating: 5
2.	REDTECH USB-C to Lightning Cable 3.3FT – Average Rating: 5
3.	Amazon Basics Wireless Mouse – Average Rating: 5


Which products have the highest number of reviews?
Based on the dataset analysis, the product with the highest reviews are as follows:
1.	AmazonBasics Flexible Premium – Rating count of 426,973
2.	AmazonBasics High Speed – Rating count of 426,973

How many products have a discount of 50% or more?
= 1465 products

How many products have fewer than 1,000 reviews?
= 326 products

Which categories have products with the highest discounts?
= Computer and Accessories

Identify the top 5 products in terms of rating and number of reviews combined.
1.	AmazonBasics Flexible Premium
2.	AmazonBasics High Speed
3.	AmazonBasics High Speed
4.	AmazonBasics Flexible Premium
5.	boAT Bassheads 100


### Overall Interpretations:

* **Electronics and Computers & Accessories Dominate:** These two categories are not only highly represented in terms of product count but also drive the majority of customer engagement (reviews) and revenue potential due to their higher price points and substantial sales volume. They also frequently feature high discount percentages, which likely contributes to their popularity.
* **Discounting as a Strategy:** Discounts are a prevalent pricing strategy across many categories, particularly in Electronics and Computers & Accessories. While heavily discounted products (91-100%) show high average ratings, which is a positive sign for customer value perception, it's interesting that the very lowest discount bracket (0-10%) also maintains high ratings. This could imply that some products are inherently popular and well-received regardless of discounts, or that strategically applied smaller discounts can still drive satisfaction.
* **Importance of Customer Reviews:** The sheer volume of rating counts underscores the significance of customer reviews in the Amazon ecosystem. Detailed review content (as seen in the raw data) provides critical qualitative insights that complement the quantitative findings from the pivot table.
* **Revenue Potential:** Categories with higher product counts, average prices, and review volumes (like Electronics and Computers & Accessories) naturally hold the largest potential for revenue. Further analysis of the "total potential revenue" metric (if calculated correctly) would confirm this.
