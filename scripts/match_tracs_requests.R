# 📌 Load required libraries ----
library(readr)
library(readxl)
library(dplyr)
library(stringr)
library(stringdist)
library(janitor)
library(lubridate)
library(tidyr)
library(writexl)
library(purrr)

# 1️⃣ Clean directory properly ----
directory$Name <- iconv(directory$Name, from = "", to = "UTF-8")

directory <- directory %>%
  mutate(
    Clean_Name = str_trim(str_to_upper(Name)),
    Clean_Name = str_remove(Clean_Name, "\\b(JR|SR|II|III|IV|V)\\b"),
    Clean_Name = str_replace_all(Clean_Name, "\\b[A-Z]\\.?\\b", ""), # Remove lone initials
    Clean_Name = str_squish(Clean_Name),
    Nickname = str_extract(Clean_Name, "\\((.*?)\\)"),
    Nickname = str_remove_all(Nickname, "[\\(\\)]") %>% str_trim(),
    Clean_Name_No_Nickname = str_remove(Clean_Name, "\\s*\\(.*?\\)"),
    Clean_Name_No_Nickname = str_squish(Clean_Name_No_Nickname),
    Clean_Name_No_Nickname = str_remove_all(Clean_Name_No_Nickname, "\\."),
    Nickname = str_remove_all(Nickname, "\\."),
    Directory_Email = `E-Mail Address`,
    Directory_Email_Username = str_split_fixed(Directory_Email, "@", 2)[,1] %>%
      str_replace_all("[[:punct:]]", " ") %>%
      str_squish() %>%
      str_to_upper(),
    Final_Names = ifelse(
      !is.na(Nickname) & Nickname != "",
      paste(Clean_Name_No_Nickname, "|", Nickname),
      Clean_Name_No_Nickname
    )
  )

# 2️⃣ Clean & filter requests ----
requests_clean <- requests %>%
  mutate(
    Submitted_Date = as.numeric(`Submitted Date`),
    Submitted_Date_Converted = as.Date(Submitted_Date, origin = "1899-12-30"),
    Closed_Date = as.numeric(`Closed Date`),
    Closed_Date_Converted = as.Date(Closed_Date, origin = "1899-12-30")
  ) %>%
  filter(
    Submitted_Date_Converted >= as.Date("2024-01-01") |
      Closed_Date_Converted >= as.Date("2024-01-01")
  )

# 3️⃣ Clean PI name and PI email ----
requests_clean <- requests_clean %>%
  mutate(
    PI_Raw = str_trim(`Project Principal Investigator`),
    PI_Email_Raw = str_trim(`PI Email`),
    Clean_PI = if_else(
      str_detect(PI_Raw, "@"),
      str_to_upper(
        str_replace_all(
          str_split_fixed(PI_Raw, "@", 2)[,1],
          "[[:punct:]]",
          " "
        )
      ),
      str_to_upper(
        str_replace_all(PI_Raw, "[[:punct:]]", "")
      )
    ),
    Clean_PI_Email = if_else(
      str_detect(PI_Email_Raw, "@"),
      str_to_upper(
        str_replace_all(
          str_split_fixed(PI_Email_Raw, "@", 2)[,1],
          "[[:punct:]]",
          " "
        )
      ),
      NA_character_
    )
  )

# 4️⃣ Improved nickname map ----
nickname_map <- c(
  "JAKE" = "JACOB",
  "JACK" = "JACOB",
  "BOB" = "ROBERT",
  "BILL" = "WILLIAM",
  "JIM" = "JAMES",
  "KATE" = "KATHERINE",
  "STEVE" = "STEVEN",
  "SUE" = "SUSAN",
  "CHRIS" = "CHRISTOPHER",
  "KIM" = "KIMBERLY"
)

reverse_nick <- function(name) {
  revs <- names(nickname_map)[nickname_map == name]
  if (length(revs) > 0) revs[1] else NA_character_
}

requests_clean <- requests_clean %>%
  mutate(
    PI_First = word(Clean_PI, 1),
    PI_First_Nick = if_else(
      PI_First %in% names(nickname_map),
      nickname_map[PI_First],
      NA_character_
    ),
    PI_First_Alias = map_chr(PI_First, reverse_nick),
    Clean_PI_Nickname = if_else(
      !is.na(PI_First_Nick),
      str_replace(Clean_PI, PI_First, PI_First_Nick),
      Clean_PI
    ),
    Clean_PI_Alias = if_else(
      !is.na(PI_First_Alias),
      str_replace(Clean_PI, PI_First, PI_First_Alias),
      Clean_PI
    )
  )

# 5️⃣ Create match keys ----
requests_keys <- requests_clean %>%
  mutate(Row_ID = row_number()) %>%
  select(Row_ID, Clean_PI, Clean_PI_Nickname, Clean_PI_Alias, Clean_PI_Email) %>%
  pivot_longer(
    cols = c(Clean_PI, Clean_PI_Nickname, Clean_PI_Alias, Clean_PI_Email),
    names_to = "Key_Type",
    values_to = "Key_Value"
  ) %>%
  filter(!is.na(Key_Value))

# 6️⃣ Expand directory nicknames & no-middle-initial key ----
directory_keys <- directory %>%
  mutate(
    Directory_ID = row_number(),
    Final_Name1 = str_split_fixed(Final_Names, "\\|", 2)[, 1] %>% str_trim(),
    Final_Name2 = str_split_fixed(Final_Names, "\\|", 2)[, 2] %>% str_trim(),
    Final_Name1_NoMI = str_remove_all(Final_Name1, "\\b[A-Z]\\b") %>% str_squish(),
    Final_Name1_First = word(Final_Name1, 1),
    Final_Name1_Alias = map_chr(Final_Name1_First, reverse_nick)
  ) %>%
  select(
    Directory_ID,
    Final_Name1, Final_Name1_NoMI, Final_Name2, Final_Name1_Alias,
    Directory_Email_Username
  ) %>%
  pivot_longer(
    cols = c(Final_Name1, Final_Name1_NoMI, Final_Name2, Final_Name1_Alias, Directory_Email_Username),
    names_to = "Key_Type",
    values_to = "Directory_Key"
  ) %>%
  filter(!is.na(Directory_Key))

# 7️⃣ Cross join & fuzzy match ----
fuzzy_matches <- requests_keys %>%
  inner_join(directory_keys, by = character()) %>%
  mutate(
    dist = stringdist(Key_Value, Directory_Key, method = "lv"),
    max_dist = ceiling(0.2 * pmax(str_length(Key_Value), str_length(Directory_Key)))
  ) %>%
  filter(dist <= max_dist)

# 8️⃣ Final matched output ----
matched_requests <- fuzzy_matches %>%
  distinct(Row_ID) %>%
  inner_join(
    requests_clean %>% mutate(Row_ID = row_number()),
    by = "Row_ID"
  )

# 9️⃣ Export to Excel ----
writexl::write_xlsx(matched_requests, "matched_requests.xlsx")
