# ---------------------------------------------------------------------------------------------------------
# Process Total Population (1970-2024) and Japanese Population(2015-2024) by Prefecture (1970-2024)
# Author: Tshewang Gyeltshen, GHP, GSM, The University of Tokyo
# Source File Links: 
#   - Download raw data for 2007 - 2024 from:
#     - https://www.e-stat.go.jp/en/stat-search/files?page=1&layout=datalist&toukei=00200524&bunya_l=02&tstat=000000090001&cycle=7&tclass1=000001011679&result_back=1&tclass2val=0&metadata=1&data=1
#   - 1970 - 2000:
#     - https://www.e-stat.go.jp/stat-search/file-download?statInfId=000000090269&fileKind=0
#   - 2000 - 2020:
#     - https://www.e-stat.go.jp/stat-search/file-download?statInfId=000013168609&fileKind=4
#     - https://www.e-stat.go.jp/stat-search/file-download?statInfId=000032212354&fileKind=4
# Notes:
#   - Extracts population by prefecture, sex (1=male, 2=female), age group, and year
#   - Age Group (1970-2024), and Sex (2007–2024); 
#   - Total Population (1970-2024); Japanese Population(2015-2024)
# ---------------------------------------------------------------------------------------------------------
rm(list = ls())
library(readxl)
library(dplyr)
library(tidyr)
library(stringr)
library(purrr)
# ---------------------------------------------------------------------------------------------------------
# Process 1970–2000 Total Population by Prefecture and Age Group
# Note: Age groups vary slightly by year. No sex-specific info in this dataset.
# ---------------------------------------------------------------------------------------------------------

# File path
file_path <- "C:/Users/kench/Downloads/estat_pop/total_agegroup_1970_2024/poptotal_1970_2000.xlsx"
pref_endings <- "-(ken|fu|to|do)$"

# Japanese-to-Gregorian year mapping
sheet_years <- c(
        "昭和45年" = 1970, "昭和46年" = 1971, "昭和47年" = 1972, "昭和48年" = 1973,
        "昭和49年" = 1974, "昭和50年" = 1975, "昭和51年" = 1976, "昭和52年" = 1977,
        "昭和53年" = 1978, "昭和54年" = 1979, "昭和55年" = 1980, "昭和56年" = 1981,
        "昭和57年" = 1982, "昭和58年" = 1983, "昭和59年" = 1984, "昭和60年" = 1985,
        "昭和61年" = 1986, "昭和62年" = 1987, "昭和63年" = 1988, "平成元年" = 1989,
        "平成2年" = 1990, "平成3年" = 1991, "平成4年" = 1992, "平成5年" = 1993,
        "平成6年" = 1994, "平成7年" = 1995, "平成8年" = 1996, "平成9年" = 1997,
        "平成10年" = 1998, "平成11年" = 1999, "平成12年" = 2000
)
# Unified column names
new_colnames <- c("prefno", "prefname1", "prefname2", "totalpop", 
                  "age_0_4", "age_5_9", "age_10_14", "age_15_19", "age_20_24", 
                  "age_25_29", "age_30_34", "age_35_39", "age_40_44", "age_45_49", 
                  "age_50_54", "age_55_59", "age_60_64", "age_65_69", "age_70_74", 
                  "age_75_79", "age_80_84", "age_85_plus", "row", "sex", "year")

# Function to clean a sheet
process_sheet1970_2000 <- function(year, sheet_name) {
        df <- read_excel(file_path, sheet = sheet_name, skip = 7)
        
        # Remove trailing rows and the first empty column
        df <- df %>% filter(!row_number() %in% c(1:4, 52:58))
        df <- df[, -1]
        # Assign column names
        # Ensure df has the correct number of columns to match new_colnames
        if (ncol(df) < length(new_colnames)) {
                # Calculate how many columns are missing
                missing_cols <- length(new_colnames) - ncol(df)
                
                # Add NA columns with temporary names
                for (i in 1:missing_cols) {
                        df[[paste0("temp_col_", i)]] <- NA
                }
        }
        
        # Now safely assign new column names
        colnames(df) <- new_colnames
        df$prefname2 <- str_replace(df$prefname2, pref_endings, "" )
        
        #assign year
        df$year <- year
        # Coerce all age columns to numeric
        df <- df %>%
                mutate(across(starts_with("age_"), ~ suppressWarnings(as.numeric(.))))
        
        df <- df %>%
                pivot_longer(cols = starts_with("age_"), 
                             names_to = "age_group", 
                             values_to = "population") %>%
                select(prefName = prefname2, year, age_group, population)
        
        return(df)
}

# Process all sheets
pop1970_2000 <- imap_dfr(sheet_years, process_sheet1970_2000)

# Preview cleaned data
head(pop1970_2000)

# Optional: write to CSV
write.csv(pop1970_2000, "total_agegroup_1970_2024/total_population_by_agegrp_1970_2000.csv", row.names = FALSE)


# ---------------------------------------------------------------------------------------------------------
# Process 2000–2020 Total Population by Prefecture, Age Group, and Sex 
# Notes:
#   - Handles multi-format sheets: 平成-based (2000–2014) and Gregorian (2015–2020)
#   - Extracts Total Population (Foreigners + Japanese) by prefecture, sex (1=male, 2=female), and age group
# ---------------------------------------------------------------------------------------------------------

# File path
file_path <- "C:/Users/kench/Downloads/estat_pop/total_agegroup_1970_2024/poptotal_2000_2020.xlsx"

# Regex to remove suffixes like -ken, -fu, -to, -do
pref_endings <- "-(ken|fu|to|do)$"

# Common column names based on sheet structure
new_colnames <- c(
        "prefno", "prefname1", "prefname2", "totalpop", 
        "age_0_4", "age_5_9", "age_10_14", "age_15_19", "age_20_24", 
        "age_25_29", "age_30_34", "age_35_39", "age_40_44", "age_45_49", 
        "age_50_54", "age_55_59", "age_60_64", "age_65_69", "age_70_74", 
        "age_75_79", "age_80_84", "age_85_plus", "row", "sex", "year"
)

# -------------------------------------------------------------------------------------
# Helper function to generate sheet name from year
# -------------------------------------------------------------------------------------
get_sheet_name <- function(yr) {
        if (yr >= 2000 & yr <= 2014) {
                sprintf("第9表（平成%d年）", yr - 1988)
        } else if (yr >= 2015 & yr <= 2020) {
                sprintf("総人口（%d年）", yr)
        } else {
                NA_character_
        }
}

# -------------------------------------------------------------------------------------
# Main sheet processing function
# -------------------------------------------------------------------------------------
process_sheet2000_2020 <- function(yr) {
        sheet_name <- get_sheet_name(yr)
        if (is.na(sheet_name)) return(NULL)
        
        # Read sheet (skip header and metadata rows)
        df <- read_excel(file_path, sheet = sheet_name, skip = 8, col_names = TRUE) %>%
                select(-1)  # Drop first unnamed column
        
        # Pad missing columns if needed
        if (ncol(df) < length(new_colnames)) {
                missing_cols <- length(new_colnames) - ncol(df)
                for (i in 1:missing_cols) {
                        df[[paste0("temp_col_", i)]] <- NA
                }
        }
        
        # Assign standard column names
        colnames(df) <- new_colnames
        
        # Add row numbers and identify sex by known row ranges
        df <- df %>%
                mutate(
                        row = row_number(),
                        sex = case_when(
                                row >= 71 & row <= 118 ~ 1,   # Male
                                row >= 138 & row <= 185 ~ 2,  # Female
                                TRUE ~ NA_real_
                        ),
                        year = yr
                ) %>%
                filter(!is.na(sex))  # Keep only male/female rows
        
        # Clean prefecture name suffix
        df$prefname2 <- str_replace(df$prefname2, pref_endings, "")
        
        # Convert age group columns to numeric
        df <- df %>%
                mutate(across(starts_with("age_"), ~ suppressWarnings(as.numeric(.))))
        
        # Reshape to long format
        df_long <- df %>%
                pivot_longer(
                        cols = starts_with("age_"), 
                        names_to = "age_group", 
                        values_to = "population"
                ) %>%
                select(prefName = prefname2, year, sex, age_group, population)
        
        return(df_long)
}

# -------------------------------------------------------------------------------------
# Process all sheets from 2000 to 2020
# -------------------------------------------------------------------------------------
years <- 2000:2020
pop2000_2020_sex <- map_dfr(years, process_sheet2000_2020)

# -------------------------------------------------------------------------------------
# Preview result and optionally export
# -------------------------------------------------------------------------------------
head(pop2000_2020_sex)
# Export cleaned dataset
write.csv(pop2000_2020_sex, "total_agegroup_1970_2024/total_population_by_agegroup_sex_2000_2020.csv", row.names = FALSE)
# -------------------------------------------------------------------------------------
# Summarize Total Population by Prefecture, Year, and Age Group (Male + Female)/No sex
# -------------------------------------------------------------------------------------
pop2000_2020 <- pop2000_2020_sex %>%
        group_by(prefName, year, age_group) %>%
        summarise(population = sum(population, na.rm = TRUE), .groups = "drop")

# Optional: write to CSV
write.csv(pop2000_2020, "total_agegroup_1970_2024/total_population_by_agegroup_2000_2020.csv", row.names = FALSE)


# ---------------------------------------------------------------------------------------------------------
# Process 2007–2024 Total Population and Japanese Population by Prefecture, Age Group, and Sex
# Notes:
#   - Download raw data from:
#     - https://www.e-stat.go.jp/en/stat-search/files?page=1&layout=datalist&toukei=00200524&bunya_l=02&tstat=000000090001&cycle=7&tclass1=000001011679&result_back=1&tclass2val=0&metadata=1&data=1
#   - Extracts population by prefecture, sex (1=male, 2=female), age group, and year
# ---------------------------------------------------------------------------------------------------------
# Parameters
years <- 2007:2024
pref_endings <- "-(ken|fu|to|do)$"

# Standardized column names
new_colnames <- c("prefno", "prefname1", "prefname2", "totalpop", 
                  "age_0_4", "age_5_9", "age_10_14", "age_15_19", "age_20_24", 
                  "age_25_29", "age_30_34", "age_35_39", "age_40_44", "age_45_49", 
                  "age_50_54", "age_55_59", "age_60_64", "age_65_69", "age_70_74", 
                  "age_75_79", "age_80_84", "age_85_plus", "row", "sex", "year")

# Result lists
total_list <- list()
japanese_list <- list()

# Iterate through years
for (yr in years) {
        
        file_path <- paste0("C:/Users/kench/Downloads/estat_pop/pop", yr, ".xls")
        
        if (yr %in% c(2020, 2015, 2010)) {
                
                # Special years (Census years using '第２表')
                df_total <- read_excel(file_path, sheet = "第２表") %>%
                        mutate(row = row_number(),
                               sex = case_when(
                                       row %in% 150:197 ~ 1,  # Male
                                       row %in% 77:130  ~ 2,  # Female
                                       TRUE ~ NA_real_
                               )) %>%
                        filter(!is.na(sex) | row == 77) %>%
                        select(-c(1, 4)) %>%
                        mutate(year = yr)
                
                df_jap <- read_excel(file_path, sheet = "第２表") %>%
                        mutate(row = row_number(),
                               sex = case_when(
                                       row %in% 351:398 ~ 1,  # Male
                                       row %in% 278:331 ~ 2,  # Female
                                       TRUE ~ NA_real_
                               )) %>%
                        filter(!is.na(sex)) %>%
                        select(-c(1, 4)) %>%
                        mutate(year = yr)
                
        } else {
                
                # Standard years (using '第10表')
                df_total <- read_excel(file_path, sheet = "第10表") %>%
                        mutate(row = row_number(),
                               sex = case_when(
                                       row %in% 153:200 ~ 1,  # Male
                                       row %in% 86:133  ~ 2,  # Female
                                       TRUE ~ NA_real_
                               )) %>%
                        filter(!is.na(sex) | row == 80) %>%
                        select(-c(1:8, 11)) %>%
                        mutate(year = yr)
                
                df_jap <- read_excel(file_path, sheet = "第10表") %>%
                        mutate(row = row_number(),
                               sex = case_when(
                                       row %in% 354:401 ~ 1,  # Male
                                       row %in% 282:334 ~ 2,  # Female
                                       TRUE ~ NA_real_
                               )) %>%
                        filter(!is.na(sex) | row == 80) %>%
                        select(-c(1:8, 11)) %>%
                        mutate(year = yr)
        }
        
        # Assign column names
        if (ncol(df_total) == length(new_colnames)) colnames(df_total) <- new_colnames
        if (ncol(df_jap) == length(new_colnames)) colnames(df_jap) <- new_colnames
        
        # Clean prefecture names
        df_total <- df_total %>%
                filter(!row_number() %in% 1:6) %>%
                mutate(prefname2 = str_replace(prefname2, pref_endings, ""))
        
        df_jap <- df_jap %>%
                filter(!row_number() %in% 1:6) %>%
                mutate(prefname2 = str_replace(prefname2, pref_endings, ""))
        
        # Optional: Save wide format
        write.csv(df_total, paste0("totalPop_pref_", yr, ".csv"))
        write.csv(df_jap, paste0("japanese_Pop_pref_", yr, ".csv"))
        
        # Convert to long format
        df_total_long <- df_total %>%
                pivot_longer(cols = starts_with("age_"), names_to = "age_group", values_to = "population") %>%
                mutate(prefName = coalesce(prefname2, prefname1)) %>%
                select(prefName, sex, year, age_group, population)
        
        df_jap_long <- df_jap %>%
                pivot_longer(cols = starts_with("age_"), names_to = "age_group", values_to = "population") %>%
                mutate(prefName = coalesce(prefname2, prefname1)) %>%
                select(prefName, sex, year, age_group, population)
        
        # Store
        total_list[[as.character(yr)]]    <- df_total_long
        japanese_list[[as.character(yr)]] <- df_jap_long
}

# Combine all years
poptotal2007_2024    <- bind_rows(total_list)%>%
        mutate(population = as.numeric(population)) %>%
        rename(totalpop = population)
popjp2007_2024 <- bind_rows(japanese_list)%>%
        mutate(population = as.numeric(population)) %>%
        rename(jpop = population)


# Save cleaned data
write.csv(poptotal2007_2024,    "total_agegroup_1970_2024/total_population_2007_2024.csv", row.names = FALSE)
write.csv(popjp2007_2024, "total_agegroup_1970_2024/japanese_population_2015_2024.csv", row.names = FALSE)

# Merge total and Japanese population
pop2007_2024sex <- poptotal2007_2024 %>%
        left_join(popjp2007_2024, by = c("prefName", "sex", "year", "age_group"))

# Save merged file
write.csv(pop2007_2024sex, "total_agegroup_1970_2024/merged_population_totl2007_2024_vs_jap2015_2024.csv", row.names = FALSE)

# Preview
head(pop2007_2024sex)
# -------------------------------------------------------------------------------------
# Summarize Total Population by Prefecture, Year, and Age Group (Male + Female)/NO sex
# -------------------------------------------------------------------------------------
pop2007_2024 <- pop2007_2024sex %>%
        group_by(prefName, year, age_group) %>%
        summarise(population = sum(totalpop, na.rm = TRUE),
                  jpop = sum(jpop, na.rm = TRUE), .groups = "drop")



# -------------------------------------------------------------------------------------
# Merge Datasets 1970–2000 and 2000–2020; 2007 - 2024 datasets
# -------------------------------------------------------------------------------------
pop2001_2020 <- pop2000_2020 %>% filter(prefName != "Japan", year > 2000)
pop2021_2024 <- pop2007_2024 %>% filter(prefName != "Japan", year %in% 2021:2024) %>% select(-jpop)

poptotal1970_2024 <- bind_rows(pop1970_2000, pop2001_2020, pop2021_2024) %>%
        arrange(prefName, year, age_group)
write.csv(poptotal1970_2024, "total_agegroup_1970_2024/total_population_by_agegroup_prefecture_1970_2024.csv", row.names = FALSE)

# Save as RData
prefcode <- readxl::read_xlsx("total_agegroup_1970_2024/prefecture-code.xlsx")
poptotal1970_2024 <- poptotal1970_2024 %>%
        mutate(prefName = factor(prefName, levels = prefcode$Prefecture))
poptotal1970_2024lst <- split(poptotal1970_2024, poptotal1970_2024$prefName)
save(poptotal1970_2024lst, file = "total_agegroup_1970_2024/poptotal1970_2024lst.RData")

# ---------------------------------------------------------------------------------------------------------
# Other Derived Files
# ---------------------------------------------------------------------------------------------------------
pop2000_2006_sex <- pop2000_2020_sex %>% filter(!year %in% 2007:2020)
poptotal2000_2024sex <- bind_rows(pop2000_2006_sex, pop2007_2024sex) %>%
        arrange(prefName, year, age_group)
write.csv(poptotal2000_2024sex, "total_agegroup_1970_2024/total_population_by_agegroup_sex_2000_2024.csv", row.names = FALSE)

popjp2007_2024 <- popjp2007_2024 %>% arrange(prefName, year, age_group)
write.csv(popjp2007_2024, "total_agegroup_1970_2024/jp_population_by_agegroup_sex_2007_2024.csv", row.names = FALSE)
