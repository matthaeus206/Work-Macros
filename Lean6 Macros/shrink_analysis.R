library(dplyr)
library(ggplot2)
library(readr)
library(scales)

# Load data
df <- read_csv("store_inventory_data.csv")

# Calculate shrink (defects) and total opportunities
df <- df %>%
  mutate(
    Shrink = pmax(Expected_Inventory - Actual_Inventory, 0),  # defects only
    Opportunities = Expected_Inventory                        # each stocked unit is an opportunity
  )

# Summarize by store
shrink_summary <- df %>%
  group_by(Store_ID) %>%
  summarise(
    Total_Shrink = sum(Shrink),
    Total_Opportunities = sum(Opportunities),
    DPMO = (Total_Shrink / Total_Opportunities) * 1e6,
    Sigma_Level = 1.5 + qnorm(1 - (Total_Shrink / Total_Opportunities))
  ) %>%
  arrange(desc(DPMO))

print(shrink_summary)

# Bar plot of DPMO by Store
ggplot(shrink_summary, aes(x = reorder(Store_ID, -DPMO), y = DPMO)) +
  geom_bar(stat = "identity", fill = "firebrick") +
  theme_minimal() +
  labs(title = "Defects per Million Opportunities (DPMO) by Store",
       x = "Store ID", y = "DPMO") +
  scale_y_continuous(labels = comma) +
  theme(axis.text.x = element_text(angle = 90, hjust = 1))

# Drill down: top defect categories in worst store
worst_store <- shrink_summary$Store_ID[1]

category_shrink <- df %>%
  filter(Store_ID == worst_store) %>%
  group_by(Category) %>%
  summarise(Defects = sum(Shrink)) %>%
  arrange(desc(Defects)) %>%
  top_n(5, Defects)

ggplot(category_shrink, aes(x = reorder(Category, -Defects), y = Defects)) +
  geom_col(fill = "steelblue") +
  labs(title = paste("Top Shrink Categories in Store", worst_store),
       x = "Category", y = "Defective Units") +
  theme_minimal()
