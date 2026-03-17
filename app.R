library(shiny)
library(bslib)
library(dplyr)
library(stringr)
library(readxl)
library(openxlsx)
library(DT)
library(ggplot2)
library(tidyr)
library(purrr)
library(scales)

required_cols <- c("Last Name", "First Name", "Rank", "Status", "Action List")

extract_component <- function(action_list, patterns, label) {
  vals <- rep("", length(action_list))
  for (pat in patterns) {
    hit <- !is.na(action_list) & str_detect(action_list, regex(pat, ignore_case = TRUE))
    vals[hit] <- label
  }
  vals
}

process_personnel <- function(df) {
  missing_cols <- setdiff(required_cols, names(df))
  if (length(missing_cols) > 0) {
    stop(
      paste0(
        "The uploaded file is missing these required columns: ",
        paste(missing_cols, collapse = ", ")
      )
    )
  }

  aps <- df %>%
    select(all_of(required_cols)) %>%
    mutate(
      `Action List` = as.character(`Action List`),
      PHA = extract_component(`Action List`, c("\\|PHA\\|"), "PHA"),
      PHAQ = extract_component(`Action List`, c("\\|PHAQ\\|"), "PHAQ"),
      Dental = extract_component(`Action List`, c("\\|Dental\\|"), "Dental"),
      `Military Dental` = extract_component(`Action List`, c("\\|Dental\\|"), "Military Dental"),
      HIV = extract_component(`Action List`, c("\\|HIV\\|"), "HIV"),
      `Blood Type` = extract_component(`Action List`, c("\\|Blood Type\\|"), "Blood Type"),
      RH = extract_component(`Action List`, c("\\|RH\\|"), "RH"),
      G6PD = extract_component(`Action List`, c("\\|G6PD\\|"), "G6PD"),
      SickleCell = extract_component(`Action List`, c("\\|SickleCell\\|", "\\|SickeCell\\|"), "Sickle Cell"),
      Tdap = extract_component(`Action List`, c("\\|Tdap\\|"), "Tdap"),
      Td = extract_component(`Action List`, c("\\|Td\\|"), "Td"),
      Varicella = extract_component(`Action List`, c("\\|Varicella\\|"), "Varicella"),
      HepA = extract_component(`Action List`, c("\\|Hep A\\|"), "Hep A"),
      MMR = extract_component(`Action List`, c("\\|MMR\\|"), "MMR"),
      Polio = extract_component(`Action List`, c("\\|Polio\\|"), "Polio"),
      PPD = extract_component(`Action List`, c("\\|PPD"), "PPD (Baseline)"),
      MHA = extract_component(`Action List`, c("\\|MHA\\|"), "MHA"),
      Visual = extract_component(`Action List`, c("\\|Distant", "Distant Visual Acuity"), "Distant Visual Acuity")
    )

  active_only <- aps %>% filter(is.na(Status) | trimws(Status) == "")

  PHAQ_tbl <- active_only %>%
    select(`Last Name`, `First Name`, Rank, PHAQ) %>%
    filter(PHAQ == "PHAQ")

  PHA_tbl <- active_only %>%
    select(`Last Name`, `First Name`, Rank, PHA) %>%
    filter(PHA == "PHA")

  MHA_tbl <- active_only %>%
    select(`Last Name`, `First Name`, Rank, MHA) %>%
    filter(MHA == "MHA")

  Immuni_tbl <- active_only %>%
    select(`Last Name`, `First Name`, Rank, HepA, MMR, Polio, PPD, Varicella, Tdap, Td) %>%
    filter(HepA == "Hep A" | MMR == "MMR" | Polio == "Polio" | PPD == "PPD (Baseline)" |
             Varicella == "Varicella" | Tdap == "Tdap" | Td == "Td") %>%
    rowwise() %>%
    mutate(Immunization = {
      x <- c_across(all_of(c("HepA", "MMR", "Polio", "PPD", "Varicella", "Tdap", "Td")))
      x <- trimws(x)
      x <- x[x != "" & !is.na(x)]
      if (length(x) == 0) "" else paste0("|", paste(x, collapse = "|"), "|")
    }) %>%
    ungroup() %>%
    select(`Last Name`, `First Name`, Rank, Immunization)

  Visual_tbl <- active_only %>%
    select(`Last Name`, `First Name`, Rank, Visual) %>%
    filter(Visual == "Distant Visual Acuity")

  Test_tbl <- active_only %>%
    select(`Last Name`, `First Name`, Rank, HIV, `Blood Type`, RH, SickleCell, G6PD) %>%
    filter(HIV == "HIV" | `Blood Type` == "Blood Type" | RH == "RH" |
             SickleCell == "Sickle Cell" | G6PD == "G6PD") %>%
    rowwise() %>%
    mutate(Test = {
      x <- c_across(all_of(c("HIV", "Blood Type", "RH", "SickleCell", "G6PD")))
      x <- trimws(x)
      x <- x[x != "" & !is.na(x)]
      if (length(x) == 0) "" else paste0("|", paste(x, collapse = "|"), "|")
    }) %>%
    ungroup() %>%
    select(`Last Name`, `First Name`, Rank, Test)

  export_list <- list(
    "PHA" = PHA_tbl,
    "PHAQ" = PHAQ_tbl,
    "Immunization" = Immuni_tbl,
    "Tests" = Test_tbl,
    "MHA" = MHA_tbl,
    "Visual" = Visual_tbl
  )

  component_summary <- tibble(
    Component = c(
      "PHA", "PHAQ", "MHA", "Visual",
      "Hep A", "MMR", "Polio", "PPD (Baseline)", "Varicella", "Tdap", "Td",
      "HIV", "Blood Type", "RH", "Sickle Cell", "G6PD"
    ),
    Members_Needed = c(
      nrow(PHA_tbl),
      nrow(PHAQ_tbl),
      nrow(MHA_tbl),
      nrow(Visual_tbl),
      sum(active_only$HepA == "Hep A", na.rm = TRUE),
      sum(active_only$MMR == "MMR", na.rm = TRUE),
      sum(active_only$Polio == "Polio", na.rm = TRUE),
      sum(active_only$PPD == "PPD (Baseline)", na.rm = TRUE),
      sum(active_only$Varicella == "Varicella", na.rm = TRUE),
      sum(active_only$Tdap == "Tdap", na.rm = TRUE),
      sum(active_only$Td == "Td", na.rm = TRUE),
      sum(active_only$HIV == "HIV", na.rm = TRUE),
      sum(active_only$`Blood Type` == "Blood Type", na.rm = TRUE),
      sum(active_only$RH == "RH", na.rm = TRUE),
      sum(active_only$SickleCell == "Sickle Cell", na.rm = TRUE),
      sum(active_only$G6PD == "G6PD", na.rm = TRUE)
    )
  ) %>%
    arrange(desc(Members_Needed), Component)

  member_component_long <- active_only %>%
    transmute(
      `Last Name`, `First Name`, Rank,
      PHA, PHAQ, MHA, Visual,
      `Hep A` = HepA, MMR, Polio, `PPD (Baseline)` = PPD, Varicella, Tdap, Td,
      HIV, `Blood Type`, RH, `Sickle Cell` = SickleCell, G6PD
    ) %>%
    pivot_longer(
      cols = -c(`Last Name`, `First Name`, Rank),
      names_to = "Component",
      values_to = "Flag"
    ) %>%
    filter(!is.na(Flag), Flag != "") %>%
    select(-Flag) %>%
    arrange(Component, `Last Name`, `First Name`)

  member_summary <- member_component_long %>%
    count(`Last Name`, `First Name`, Rank, name = "Components_Needed") %>%
    arrange(desc(Components_Needed), `Last Name`, `First Name`)

  stats <- list(
    total_rows = nrow(df),
    active_members = nrow(active_only),
    members_with_any_requirement = n_distinct(paste(member_component_long$`Last Name`, member_component_long$`First Name`, member_component_long$Rank)),
    total_component_flags = nrow(member_component_long),
    avg_components_per_flagged_member = if (nrow(member_summary) > 0) round(mean(member_summary$Components_Needed), 2) else 0,
    max_components_for_member = if (nrow(member_summary) > 0) max(member_summary$Components_Needed) else 0
  )

  list(
    raw = df,
    cleaned = aps,
    export = export_list,
    component_summary = component_summary,
    member_component_long = member_component_long,
    member_summary = member_summary,
    stats = stats
  )
}

ui <- page_navbar(
  title = "69th APS - Personnel IMR Processor",
  theme = bs_theme(version = 5, bootswatch = "flatly"),

  nav_panel(
    "Upload & Export",
    sidebarLayout(
      sidebarPanel(
        fileInput(
          "file_upload",
          "Upload Personnel .xlsx file",
          accept = c(".xlsx")
        ),
        textInput("export_name", "Export file name", value = "Medical IMR Due.xlsx"),
        downloadButton("download_export", "Download processed workbook", class = "btn-primary"),
        br(), br(),
        uiOutput("status_message")
      ),
      mainPanel(
        card(
          card_header("What this app does"),
          card_body(
            tags$ol(
              tags$li("Reads the uploaded personnel workbook."),
              tags$li("Builds the same output sheets as your write.xlsx export: PHA, PHAQ, Immunization, Tests, MHA, and Visual."),
              tags$li("Lets you preview each output table in the app."),
              tags$li("Summarizes how many members need each requirement and adds charts/statistics.")
            )
          )
        ),
        br(),
        card(
          card_header("Source logic from your attached R script"),
          card_body(
            HTML("This Shiny app follows the workflow in your uploaded R code and sample file, with a couple of small cleanup fixes for likely typos in the original matching logic.<br/><br/>Reference: <code>R Code - Process Personnel List.txt</code>.")
          )
        )
      )
    )
  ),

  nav_panel(
    "Dashboard",
    fluidPage(
      br(),
      layout_column_wrap(
        width = 1/3,
        value_box(title = "Uploaded Rows", value = textOutput("vb_total_rows"), theme = "primary"),
        value_box(title = "Active Members (i.e., Members who are not students)", value = textOutput("vb_active_members"), theme = "success"),
        value_box(title = "Members With Needs", value = textOutput("vb_members_with_any_requirement"), theme = "warning")
      ),
      br(),
      layout_column_wrap(
        width = 1/3,
        value_box(title = "Total Requirement Flags", value = textOutput("vb_total_component_flags"), theme = "info"),
        value_box(title = "Avg Needs / Flagged Member", value = textOutput("vb_avg_components"), theme = "secondary"),
        value_box(title = "Max Needs For One Member", value = textOutput("vb_max_components"), theme = "danger")
      ),
      br(),
      layout_columns(
        card(
          full_screen = TRUE,
          card_header("Members needed by component"),
          plotOutput("component_bar", height = 420)
        ),
        card(
          full_screen = TRUE,
          card_header("Requirement mix"),
          plotOutput("component_pie", height = 420)
        )
      ),
      br(),
      card(
        full_screen = TRUE,
        card_header("Summary table"),
        DTOutput("component_summary_table")
      )
    )
  ),

  nav_panel(
    "Output Tables",
    fluidPage(
      br(),
      selectInput(
        "table_choice",
        "Choose a generated output table",
        choices = c("PHA", "PHAQ", "Immunization", "Tests", "MHA", "Visual")
      ),
      card(
        full_screen = TRUE,
        card_header(textOutput("selected_table_title")),
        DTOutput("selected_output_table")
      )
    )
  ),

  nav_panel(
    "Member Detail",
    fluidPage(
      br(),
      layout_columns(
        card(
          full_screen = TRUE,
          card_header("Members with the most outstanding components"),
          DTOutput("member_summary_table")
        ),
        card(
          full_screen = TRUE,
          card_header("All member-component requirements"),
          DTOutput("member_detail_table")
        )
      )
    )
  )
)

server <- function(input, output, session) {
  processed_data <- reactive({
    req(input$file_upload)
    ext <- tools::file_ext(input$file_upload$name)
    validate(need(tolower(ext) == "xlsx", "Please upload an .xlsx file."))

    df <- read_excel(input$file_upload$datapath)
    process_personnel(df)
  })

  output$status_message <- renderUI({
    if (is.null(input$file_upload)) {
      div(class = "alert alert-info", "Upload an .xlsx file to begin.")
    } else {
      tryCatch({
        processed_data()
        div(class = "alert alert-success", paste("Loaded:", input$file_upload$name))
      }, error = function(e) {
        div(class = "alert alert-danger", e$message)
      })
    }
  })

  output$download_export <- downloadHandler(
    filename = function() {
      nm <- trimws(input$export_name)
      if (!grepl("\\.xlsx$", nm, ignore.case = TRUE)) nm <- paste0(nm, ".xlsx")
      nm
    },
    content = function(file) {
      result <- processed_data()
      write.xlsx(result$export, file)
    }
  )

  output$vb_total_rows <- renderText({ processed_data()$stats$total_rows })
  output$vb_active_members <- renderText({ processed_data()$stats$active_members })
  output$vb_members_with_any_requirement <- renderText({ processed_data()$stats$members_with_any_requirement })
  output$vb_total_component_flags <- renderText({ processed_data()$stats$total_component_flags })
  output$vb_avg_components <- renderText({ processed_data()$stats$avg_components_per_flagged_member })
  output$vb_max_components <- renderText({ processed_data()$stats$max_components_for_member })

  output$component_summary_table <- renderDT({
    datatable(
      processed_data()$component_summary,
      extensions = c("Buttons"),
      options = list(dom = "Bfrtip", buttons = c("copy", "csv", "excel"), pageLength = 10)
    )
  })

  output$component_bar <- renderPlot({
    summary_df <- processed_data()$component_summary %>% filter(Members_Needed > 0)
    validate(need(nrow(summary_df) > 0, "No required components found in the uploaded file."))

    ggplot(summary_df, aes(x = reorder(Component, Members_Needed), y = Members_Needed)) +
      geom_col() +
      geom_text(aes(label = Members_Needed), hjust = -0.15, size = 4) +
      coord_flip() +
      scale_y_continuous(expand = expansion(mult = c(0, 0.12))) +
      labs(x = NULL, y = "Members", title = "How many members need each component") +
      theme_minimal(base_size = 13)
  })

  output$component_pie <- renderPlot({
    summary_df <- processed_data()$component_summary %>% filter(Members_Needed > 0)
    validate(need(nrow(summary_df) > 0, "No required components found in the uploaded file."))

    ggplot(summary_df, aes(x = "", y = Members_Needed, fill = Component)) +
      geom_col(width = 1) +
      coord_polar(theta = "y") +
      labs(title = "Share of all component flags", x = NULL, y = NULL) +
      theme_void(base_size = 13) +
      theme(legend.position = "right")
  })

  output$selected_table_title <- renderText({
    paste(input$table_choice, "table preview")
  })

  output$selected_output_table <- renderDT({
    tbl <- processed_data()$export[[input$table_choice]]
    datatable(
      tbl,
      extensions = c("Buttons"),
      options = list(dom = "Bfrtip", buttons = c("copy", "csv", "excel"), pageLength = 15, scrollX = TRUE)
    )
  })

  output$member_summary_table <- renderDT({
    datatable(
      processed_data()$member_summary,
      extensions = c("Buttons"),
      options = list(dom = "Bfrtip", buttons = c("copy", "csv", "excel"), pageLength = 15, scrollX = TRUE)
    )
  })

  output$member_detail_table <- renderDT({
    datatable(
      processed_data()$member_component_long,
      extensions = c("Buttons"),
      options = list(dom = "Bfrtip", buttons = c("copy", "csv", "excel"), pageLength = 15, scrollX = TRUE)
    )
  })
}

shinyApp(ui, server)

