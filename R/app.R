#' Run the ShinyEmailAI application
#'
#' Launches the ShinyEmailAI application, which provides a UI for browsing, searching,
#' and replying to Outlook emails with AI assistance.
#'
#' @param base_dir Optional base directory where context and prompts folders are located.
#'        If NULL, uses the current working directory.
#' @param ... Additional parameters passed to shinyApp()
#' @return A Shiny application object
#' @importFrom htmltools HTML
#' @importFrom stats setNames
#' @import ellmer
#' @import shinyjs
#' @importFrom ellmer chat_claude chat_openai interpolate_file content_pdf_file
#' @export
#' @examples
#' \dontrun{
#' # Run with default settings
#' email_run_app()
#'
#' # Run with custom base directory
#' email_run_app("~/ShinyEmailAI_Data")
#' }
email_run_app <- function(base_dir = NULL, ...) {
  # If no base_dir provided, use current working directory
  if (is.null(base_dir)) {
    base_dir <- getwd()
    message("No base_dir provided, using current working directory: ", base_dir)
  } else {
    # Normalize the path
    base_dir <- normalizePath(base_dir, mustWork = FALSE)
  }

  # Check if context and prompts folders exist
  context_dir <- file.path(base_dir, "context")
  prompts_dir <- file.path(base_dir, "prompts")

  # Verify required directories exist
  if (!dir.exists(base_dir)) {
    rlang::abort(
      stringr::str_c(
        "Base directory does not exist: ", base_dir,
        "\nPlease first run: email_create_context('", base_dir, "')"
      )
    )
  }

  if (!dir.exists(context_dir)) {
    rlang::abort(
      stringr::str_c(
        "Context directory does not exist: ", context_dir,
        "\nPlease first run: email_create_context('", base_dir, "')"
      )
    )
  }

  if (!dir.exists(prompts_dir)) {
    rlang::abort(
      stringr::str_c(
        "Prompts directory does not exist: ", prompts_dir,
        "\\nPlease first run: email_create_context('", base_dir, "')"
      )
    )
  }

  # Add NULL coalescing operator
  `%||%` <- function(x, y) if (is.null(x) || length(x) == 0) y else x

  # Create a modern theme with bslib
  my_theme <- bslib::bs_theme(
    version = 5,
    bootswatch = "pulse",
    primary = "#0062cc",
    secondary = "#6c757d",
    success = "#198754",
    info = "#0dcaf0",
    warning = "#ffc107",
    danger = "#dc3545",
    base_font = bslib::font_google("Inter"),
    heading_font = bslib::font_google("Montserrat"),
    code_font = bslib::font_google("Fira Code"),
    "enable-shadows" = TRUE,
    "border-radius-base" = "0.5rem"
  )

  # Apply theme adjustments for a cleaner look
  my_theme <- bslib::bs_add_rules(
    my_theme,
    ".card { box-shadow: 0 4px 6px rgba(0, 0, 0, 0.1); }
     .nav-pills .nav-link.active { background-color: var(--bs-primary); }
     textarea { border-radius: 0.5rem; }
     .action-button { border-radius: 0.5rem; font-weight: 500; }
     div.datatables { border-radius: 0.5rem; overflow: hidden; }
     .example-tag { transition: all 0.2s ease; }
     .example-tag:hover { transform: scale(1.05); }
     .card-header { font-weight: 600; font-family: 'Montserrat', sans-serif; }
     body { font-size: 0.85rem; }
     h1, .h1 { font-size: 1.7rem; }
     h2, .h2 { font-size: 1.4rem; }
     h3, .h3 { font-size: 1.2rem; }
     h4, .h4 { font-size: 1.1rem; }
     h5, .h5 { font-size: 1rem; }
     h6, .h6 { font-size: 0.9rem; }"
  )

  # Utility: Get the user's Downloads directory (cross-platform)
  downloads_dir <- function() {
    home <- normalizePath("~", winslash = "/", mustWork = TRUE)
    sysname <- Sys.info()[["sysname"]]
    if (sysname == "Windows") {
      # Try to use USERPROFILE/Downloads, fallback to home/Downloads
      userprofile <- Sys.getenv("USERPROFILE")
      if (nzchar(userprofile)) {
        downloads <- file.path(userprofile, "Downloads")
        if (dir.exists(downloads)) {
          return(downloads)
        }
      }
      downloads <- file.path(home, "Downloads")
      if (dir.exists(downloads)) {
        return(downloads)
      }
    } else if (sysname == "Darwin" || sysname == "Linux") {
      downloads <- file.path(home, "Downloads")
      if (dir.exists(downloads)) {
        return(downloads)
      }
    }
    # Fallback: just use home directory
    return(home)
  }

  # Anthropic API key management functions
  get_api_key_file_path <- function() {
    # Creates a hidden file in the app directory
    file.path(base_dir, ".api_key")
  }

  save_api_key <- function(api_key) {
    if (is.null(api_key) || !nzchar(api_key)) {
      return(FALSE)
    }

    # Write to file
    file_path <- get_api_key_file_path()
    write(api_key, file = file_path)

    # Set environment variable for current session
    Sys.setenv(ANTHROPIC_API_KEY = api_key)

    return(TRUE)
  }

  read_api_key <- function() {
    file_path <- get_api_key_file_path()

    if (!file.exists(file_path)) {
      return(NULL)
    }

    api_key <- readLines(file_path, warn = FALSE)[1]

    if (is.null(api_key) || !nzchar(api_key)) {
      return(NULL)
    }

    # Set environment variable for current session
    Sys.setenv(ANTHROPIC_API_KEY = api_key)

    return(api_key)
  }

  is_api_key_set <- function() {
    api_key <- Sys.getenv("ANTHROPIC_API_KEY", unset = NA)
    return(!is.na(api_key) && nzchar(api_key))
  }

  # Function to create a new chat session
  new_chat <- function(body_with_header,
                       model = "gpt-4o",
                       turns = NULL,
                       prompts_dir,
                       email_samples_path) {
    system_prompt_path <- file.path(prompts_dir, "prompt-system.md")
    if (!file.exists(system_prompt_path)) {
      stop("System prompt template not found: ", system_prompt_path)
    }

    email_samples_content <- if (file.exists(email_samples_path)) {
      readr::read_file(email_samples_path)
    } else {
      "No email samples available."
    }

    system_prompt <- ellmer::interpolate_file(
      system_prompt_path,
      emailExchange   = body_with_header,
      userName        = "User",
      emailSamples    = email_samples_content,
      currentDateTime = Sys.time()
    )

    is_claude <- grepl("claude", model, ignore.case = TRUE)

    if (is_claude) {
      if (!is_api_key_set()) stop("Anthropic API key not set. Please set your API key first.")
      ellmer::chat_anthropic(
        system_prompt = system_prompt,
        model         = model,
        turns         = turns,
        api_args      = list(temperature = 1)
      )
    } else {
      ellmer::chat_openai(
        system_prompt = system_prompt,
        model         = model,
        turns         = turns,
        api_args      = list(temperature = 1)
      )
    }
  }

  # Visually enhanced minimal UI layout
  ui <- shiny::fluidPage(
    shinyjs::useShinyjs(),
    theme = my_theme,
    shiny::tags$style(shiny::HTML("
      body {
        background: linear-gradient(120deg, #f8fafc 0%, #e9f1fb 100%);
        min-height: 100vh;
      }

      /* Make DataTable more compact */
      .dataTables_wrapper .dataTables_scrollBody table.dataTable td,
      .dataTables_wrapper .dataTables_scrollBody table.dataTable th {
        padding: 5px 6px !important;
        font-size: 0.82rem;
      }
      /* Increase email preview height */
      .email-preview {
        min-height: 320px;
        background: linear-gradient(100deg, #f9fbfd 60%, #eaf3fa 100%);
        border-radius: 0.9rem;
        padding: 0.7rem 0.8rem 0.7rem 0.8rem;
        margin-bottom: 0.5rem;
        box-shadow: 0 2px 12px rgba(0,0,0,0.07);
        font-size: 0.85rem;
        transition: min-height 0.2s;
        display: flex;
        flex-direction: column;
        justify-content: flex-start;
      }
      .app-header {
        position: sticky;
        top: 0;
        z-index: 1000;
        background: rgba(255,255,255,0.85);
        backdrop-filter: blur(6px);
        box-shadow: 0 2px 12px rgba(0,0,0,0.06);
        border-radius: 0 0 1.5rem 1.5rem;
        margin-bottom: 1.5rem;
        padding: 0.8rem 1.5rem 0.7rem 1.5rem;
        display: flex;
        align-items: center;
        gap: 1rem;
      }
      .app-title {
        font-family: 'Montserrat',sans-serif;
        font-weight: 800;
        font-size: 1.4rem;
        color: #0062cc;
        letter-spacing: 0.8px;
        margin-bottom: 0;
      }
      .main-card {
        background: #fff;
        border-radius: 1.2rem;
        box-shadow: 0 8px 40px rgba(0,0,0,0.12);
        padding: 1.2rem 1.2rem 1rem 1.2rem;
        margin-bottom: 1.5rem;
        transition: box-shadow 0.2s;
        min-height: 600px;
        display: flex;
        flex-direction: column;
        justify-content: flex-start;
      }
      .main-card:hover {
        box-shadow: 0 16px 48px rgba(0,0,0,0.17);
      }
      .section-title {
        font-family: 'Montserrat',sans-serif;
        font-weight: 700;
        font-size: 0.85rem;
        margin-bottom: 0.4rem;
        color: #0062cc;
        letter-spacing: 0.4px;
        display: flex;
        align-items: center;
        gap: 0.4rem;
      }
      .btn-primary, .btn-success, .btn-outline-danger {
        margin-right: 0.5rem;
        font-size: 0.8rem;
        padding: 0.4rem 0.9rem;
        border-radius: 0.5rem;
        transition: box-shadow 0.15s, background 0.15s;
      }
      .btn-primary:hover, .btn-success:hover, .btn-outline-danger:hover {
        box-shadow: 0 2px 12px rgba(0,98,204,0.10);
        filter: brightness(1.07);
      }
      .btn-outline-danger { border-width: 2px; }
      .draft-actions { margin-top: 0.4rem; }
      .shiny-input-container { margin-bottom: 0.4rem; }
      .card { border-radius: 1rem; box-shadow: 0 2px 8px rgba(0,0,0,0.06); }
      .card-header { background: #f5f7fa; border-bottom: 1px solid #e9ecef; border-radius: 1rem 1rem 0 0; font-weight: 600; }
      .card-body { padding: 1rem; }
      .draft-box {
        background: linear-gradient(100deg, #f9fbfd 60%, #eaf3fa 100%);
        border-radius: 0.9rem;
        padding: 0.9rem;
        margin-bottom: 0.9rem;
        box-shadow: 0 2px 8px rgba(0,0,0,0.04);
      }
      .draft-box textarea, .email-preview textarea {
        font-size: 0.85rem !important;
        min-height: 70px;
      }
      .inbox-table-area {
        height: 600px;
        overflow-y: auto;
        margin-top: 1rem;
        border-radius: 1rem;
        background: linear-gradient(100deg, #f9fbfd 60%, #eaf3fa 100%);
        box-shadow: 0 2px 12px rgba(0,0,0,0.06);
        padding: 0.8rem 0.6rem;
        display: flex;
        flex-direction: column;
        justify-content: flex-start;
      }
      @media (max-width: 991px) {
        .main-card { padding: 1.2rem 0.7rem 1rem 0.7rem; min-height: 400px; }
        .app-header { padding: 1rem 0.7rem 0.7rem 0.7rem; }
        .email-preview { min-height: 260px; padding: 1.2rem 0.7rem 1.2rem 0.7rem; }
        .inbox-table-area { min-height: 220px; padding: 0.7rem 0.3rem; }
      }
      @media (max-width: 767px) {
        .main-card, .app-header { border-radius: 0.7rem; }
        .section-title { font-size: 1.1rem; }
        .app-title { font-size: 1.3rem; }
        .email-preview { min-height: 120px; }
        .inbox-table-area { min-height: 120px; }
      }
    ")),
    shiny::div(
      class = "app-header",
      shiny::icon("envelope-open-text", class = "text-primary", style = "font-size: 1.8rem;"),
      shiny::span(class = "app-title", "ShinyEmailAI"),
      shiny::div(
        style = "margin-left: auto; display: flex; gap: 10px;",
        shiny::actionButton(
          "api_key_btn",
          "Set Anthropic API Key",
          icon = shiny::icon("lock"),
          class = "btn-success"
        ),
        shiny::actionButton(
          "auth_btn",
          "Authenticate Outlook",
          icon = shiny::icon("key"),
          class = "btn-primary"
        )
      )
    ),
    shiny::fluidRow(
      shiny::column(
        width = 5,
        shiny::div(
          class = "main-card",
          shiny::div(class = "section-title", shiny::icon("inbox"), "Inbox"),
          shiny::div(
            style = "display: flex; gap: 0.5rem; margin-bottom: 0.6rem;",
            shiny::actionButton(
              "refresh_btn",
              "Refresh",
              icon = shiny::icon("sync"),
              class = "btn-primary",
              style = "opacity: 0.65;", # Start with a disabled look
              disabled = "disabled" # Start disabled
            ),
            shiny::dateRangeInput(
              "date_range",
              NULL,
              start = Sys.Date() - 30,
              end = Sys.Date(),
              separator = " to ",
              width = "100%"
            )
          ),
          shiny::textInput(
            "search_text",
            NULL,
            placeholder = "Search emails...",
            width = "100%"
          ),
          shiny::selectInput(
            "sort_order",
            NULL,
            choices = c(
              "Date (Newest First)" = "date_desc",
              "Date (Oldest First)" = "date_asc",
              "Sender (A-Z)" = "sender_asc",
              "Subject (A-Z)" = "subject_asc"
            ),
            width = "100%"
          ),
          shiny::div(
            class = "inbox-table-area",
            DT::DTOutput("email_table")
          )
        )
      ),
      shiny::column(
        width = 7,
        shiny::div(
          class = "main-card",
          shiny::div(class = "section-title", shiny::icon("envelope"), "Email Preview"),
          shiny::div(
            class = "email-preview",
            shiny::uiOutput("email_body")
          ),
          shiny::hr(style = "margin: 0.5rem 0 0.5rem 0;"),
          shiny::div(class = "section-title", shiny::icon("magic"), "Draft & Reply"),
          shiny::fluidRow(
            shiny::column(
              width = 4,
              shiny::selectInput(
                "model_choice",
                label = NULL,
                choices = c(
                  # "GPT-4o" = "gpt-4o",
                  # "GPT-4.1" = "gpt-4.1",
                  # "GPT-4.1-mini" = "gpt-4.1-mini",
                  # "o4-mini" = "o4-mini",
                  # "o3-mini" = "o3-mini",
                  "Claude 4 Sonnet" = "claude-sonnet-4-20250514",
                  "Claude 4 Opus" = "claude-opus-4-20250514",
                  "Claude 3.7 Sonnet" = "claude-3-7-sonnet-latest",
                  "Claude 3.5 Sonnet" = "claude-3-5-sonnet-latest",
                  "Claude 3.5 Haiku" = "claude-3-5-haiku-latest"
                ),
                selected = "claude-3-7-sonnet-latest",
                width = "100%"
              )
            ),
            shiny::column(
              width = 4,
              shiny::selectInput(
                "context_choice",
                label = NULL,
                choices = c("No context" = ""),
                selected = "",
                width = "100%"
              )
            ),
            shiny::column(
              width = 4,
              shiny::actionButton(
                "iterate_btn",
                "Iterate Draft",
                icon = shiny::icon("refresh"),
                class = "btn-primary w-100"
              )
            )
          ),
          shiny::div(
            class = "draft-box",
            style = "margin-top: 0.4rem;",
            shiny::textAreaInput(
              "user_request",
              label = NULL,
              value = "Craft an appropriate reply",
              width = "100%",
              height = "50px"
            )
          ),
          shiny::div(
            class = "draft-box",
            style = "margin-top: 0.4rem;",
            shiny::textAreaInput(
              "response_text",
              label = NULL,
              value = "",
              width = "100%",
              height = "80px"
            ),
            shiny::div(
              class = "draft-actions",
              shiny::actionButton(
                "clear_draft_btn",
                "Clear Draft",
                icon = shiny::icon("trash"),
                class = "btn-outline-danger"
              ),
              shiny::actionButton(
                "create_reply_draft",
                "'Reply' draft",
                icon = shiny::icon("reply"),
                class = "btn-success"
              ),
              shiny::actionButton(
                "create_reply_all_draft",
                "'Reply All' draft",
                icon = shiny::icon("reply-all"),
                class = "btn-success"
              )
            )
          )
        )
      )
    )
  )

  server <- function(input, output, session) {
    # Show privacy notice on startup
    shiny::showModal(shiny::modalDialog(
      title = "Privacy Notice",
      shiny::tagList(
        shiny::p(
          "This application does not store or log any data from your emails or interactions.",
          "All data is processed in-memory and is discarded when you close the application."
        ),
        shiny::p(
          "This application uses Anthropic's Claude AI model. By using this application, you acknowledge that you have read and understood:",
          shiny::tags$ul(
            shiny::tags$li(
              shiny::tags$a(
                href = "https://www.anthropic.com/legal/privacy",
                "Anthropic's Privacy Policy",
                target = "_blank"
              )
            ),
            shiny::tags$li(
              shiny::tags$a(
                href = "https://privacy.anthropic.com/en/articles/10023580-is-my-data-used-for-model-training",
                "Anthropic's Data Usage FAQ",
                target = "_blank"
              )
            )
          )
        )
      ),
      footer = shiny::tagList(
        shiny::actionButton(
          "privacy_accept",
          "I Acknowledge",
          class = "btn-primary"
        )
      ),
      size = "m",
      easyClose = FALSE
    ))

    # Only initialize the app after user acknowledges privacy notice
    shiny::observeEvent(input$privacy_accept,
      {
        shiny::removeModal()
      },
      once = TRUE
    )

    # Single combined reactiveValues
    values <- shiny::reactiveValues(
      outlook = NULL,
      emails_raw = NULL,
      emails_df = NULL,
      selected_email_body = NULL,
      selected_email_id = NULL,
      selected_body_with_header = NULL,
      chat_store = list(),
      context_files = NULL,
      api_key = NULL
    )

    # Load API key at startup
    shiny::observe({
      # Try to load API key from file
      api_key <- read_api_key()
      if (!is.null(api_key)) {
        values$api_key <- api_key
        shiny::showNotification("Anthropic API key loaded", type = "message")
      }
    })

    # Populate context folders on startup
    shiny::observe({
      # Start with "No context" option
      choices <- c("No context" = "")

      # Add folders if they exist
      if (dir.exists(context_dir)) {
        context_folders <- list.dirs(context_dir, recursive = FALSE, full.names = FALSE)
        if (length(context_folders) > 0) {
          # Create named vector for folders
          folder_choices <- setNames(context_folders, context_folders)
          # Combine with "No context" while preserving names
          choices <- c(choices, folder_choices)
        }
      }

      # Update the select input with all choices
      shiny::updateSelectInput(
        session,
        "context_choice",
        choices = choices,
        selected = "" # Default to "No context"
      )
    })

    # Update context files when context choice changes
    shiny::observeEvent(input$context_choice, {
      if (input$context_choice == "") {
        values$context_files <- NULL
      } else {
        values$context_files <- dir(
          file.path(context_dir, input$context_choice),
          pattern = "*.pdf",
          full.names = TRUE
        )
      }

      # Reset chat but keep draft if there's a selected email
      if (!is.null(values$selected_email_id)) {
        # Store current draft
        current_draft <- values$chat_store[[values$selected_email_id]]$draft

        # Reset chat
        values$chat_store[[values$selected_email_id]] <- list(
          chat = NULL,
          draft = current_draft, # Keep the current draft
          model = NULL
        )
      }
    })

    # Handle API key setting
    shiny::observeEvent(input$api_key_btn, {
      # Create modal dialog for API key input
      shiny::showModal(shiny::modalDialog(
        title = "Set Anthropic API Key",
        shiny::textInput("api_key_input", "Enter your Anthropic API Key:",
          value = values$api_key %||% "",
          width = "100%",
          placeholder = "sk-ant-api..."
        ),
        footer = shiny::tagList(
          shiny::modalButton("Cancel"),
          shiny::actionButton("save_api_key", "Save", class = "btn-success")
        ),
        size = "m",
        easyClose = TRUE
      ))
    })

    # Handle saving the API key
    shiny::observeEvent(input$save_api_key, {
      shiny::req(input$api_key_input)

      api_key <- trimws(input$api_key_input)

      if (nzchar(api_key)) {
        # Show a notification
        id <- shiny::showNotification("Saving API key...", type = "message", duration = NULL)

        tryCatch(
          {
            # Save to file and set in environment
            success <- save_api_key(api_key)

            # Store in reactive values
            values$api_key <- api_key

            # Enable/disable refresh button based on inbox data
            shiny::observe({
              if (!is.null(values$emails_df)) {
                shinyjs::enable("refresh_btn")
                shinyjs::runjs("$('#refresh_btn').css('opacity', '1');")
              } else {
                shinyjs::disable("refresh_btn")
                shinyjs::runjs("$('#refresh_btn').css('opacity', '0.65');")
              }
            })

            # Close modal
            shiny::removeModal()

            # Update notification
            shiny::removeNotification(id)
            shiny::showNotification("API key saved and activated successfully", type = "message", duration = 5)
          },
          error = function(e) {
            shiny::removeNotification(id)
            shiny::showNotification(paste("Error saving API key:", e$message), type = "error", duration = 10)
          }
        )
      } else {
        shiny::showNotification("Please enter a valid API key", type = "warning", duration = 5)
      }
    })

    # Initialize on startup - now optional with manual auth
    shiny::observeEvent(input$auth_btn, {
      # Check for API key first
      if (!is_api_key_set()) {
        shiny::showNotification("Please set your Anthropic API key first", type = "warning", duration = 5)
        return()
      }

      # Show modal dialog for outlook type selection
      shiny::showModal(shiny::modalDialog(
        title = "Choose Outlook Type",
        shiny::div(
          style = "display: flex; gap: 10px; justify-content: center;",
          shiny::actionButton(
            "auth_business",
            "Business Outlook",
            icon = shiny::icon("building"),
            class = "btn-primary",
            style = "width: 160px; font-weight: 500;"
          ),
          shiny::actionButton(
            "auth_personal",
            "Personal Outlook",
            icon = shiny::icon("user"),
            class = "btn-primary",
            style = "width: 160px; font-weight: 500;"
          )
        ),
        footer = shiny::modalButton("Cancel"),
        size = "s",
        easyClose = TRUE
      ))
    })

    # Handle outlook authentication for both business and personal
    shiny::observeEvent(input[["auth_business"]], {
      handle_outlook_auth("business")
    })

    shiny::observeEvent(input[["auth_personal"]], {
      handle_outlook_auth("personal")
    })

    # Helper function to handle authentication
    handle_outlook_auth <- function(type) {
      shiny::removeModal()
      shiny::showNotification("Authenticating with Outlook... You may need to check your R console for a URL.", type = "message", duration = NULL, id = "auth_status")

      # Set browser options to ensure automatic browser opening
      options(
        azure_storage_use_browser = TRUE,
        microsoft.graph.use_browser = TRUE,
        AzureR.use_browser = TRUE
      )

      # Use utils::browseURL for more reliable browser opening
      options(browser = function(url) {
        message("Opening authentication URL in browser...")
        utils::browseURL(url)
      })

      tryCatch(
        {
          # Use the appropriate function based on type
          values$outlook <- if (type == "business") {
            Microsoft365R::get_business_outlook()
          } else {
            Microsoft365R::get_personal_outlook()
          }

          values$emails_raw <- values$outlook$list_emails()
          values$emails_df <- purrr::map_dfr(
            values$emails_raw,
            ~ tidy_email(.x, include_body = "html")
          )

          shiny::removeNotification(id = "auth_status")
          shiny::showNotification(
            sprintf("Successfully authenticated with %s Outlook!", type),
            type = "message",
            duration = 5
          )
        },
        error = function(e) {
          shiny::removeNotification(id = "auth_status")
          shiny::showNotification(paste("Authentication error:", e$message), type = "error", duration = 10)
        }
      )
    }

    # Filter and sort emails based on user inputs
    filtered_emails <- shiny::reactive({
      shiny::req(values$emails_df)
      df <- values$emails_df

      # Apply date filter if available
      if (!is.null(input$date_range)) {
        date_start <- as.Date(input$date_range[1])
        date_end <- as.Date(input$date_range[2])

        df <- dplyr::filter(
          df,
          as.Date(sent_datetime) >= date_start &
            as.Date(sent_datetime) <= date_end
        )
      }

      # Apply text search if available
      if (nzchar(input$search_text)) {
        pattern <- tolower(input$search_text)
        df <- dplyr::filter(
          df,
          stringr::str_detect(tolower(subject), stringr::fixed(pattern)) |
            stringr::str_detect(tolower(from_names), stringr::fixed(pattern))
        )
      }

      # Sort the data
      if (!is.null(input$sort_order)) {
        df <- switch(input$sort_order,
          "date_desc" = dplyr::arrange(df, dplyr::desc(sent_datetime)),
          "date_asc" = dplyr::arrange(df, sent_datetime),
          "sender_asc" = dplyr::arrange(df, from_names),
          "subject_asc" = dplyr::arrange(df, subject),
          df # default, no change
        )
      }

      return(df)
    })

    # Refresh inbox when button is clicked
    shiny::observeEvent(input$refresh_btn, {
      # Check if authenticated first
      if (is.null(values$outlook)) {
        shiny::showNotification(
          "Please authenticate with Outlook before refreshing",
          type = "warning",
          duration = 5
        )
        return()
      }

      shiny::showNotification("Refreshing inbox...", type = "message", duration = NULL)

      tryCatch(
        {
          # Refresh the emails
          values$emails_raw <- values$outlook$list_emails()
          values$emails_df <- purrr::map_dfr(
            values$emails_raw,
            ~ tidy_email(.x, include_body = "html")
          )

          shiny::showNotification("Inbox refreshed successfully", type = "message")
        },
        error = function(e) {
          shiny::showNotification(
            paste("Error refreshing inbox:", e$message),
            type = "error",
            duration = 10
          )
        }
      )
    })

    # Render the DataTable *without* the id column
    output$email_table <- DT::renderDT({
      # Check if authenticated
      if (is.null(values$outlook)) {
        return(DT::datatable(
          data.frame(Message = "Please authenticate with Outlook using the button above"),
          options = list(
            dom = "t",
            ordering = FALSE,
            paging = FALSE
          ),
          selection = "none",
          rownames = FALSE
        ))
      }

      display_df <- dplyr::select(
        filtered_emails(),
        `Date/Time` = sent_datetime,
        Subject = subject,
        From = from_names
      ) # id column removed

      # Format the table
      DT::datatable(
        display_df,
        selection = "single",
        options = list(
          pageLength = 50,
          scrollY = "500px", # Fixed height for table
          scrollCollapse = TRUE,
          scrollX = TRUE,
          order = list(0, "desc"), # order by first column (date) descending
          dom = "rti", # 'r' for processing, 't' for table, 'i' for info
          language = list(
            info = "_TOTAL_ emails",
            infoEmpty = "0 emails",
            infoFiltered = "(_MAX_ total)"
          ),
          autoWidth = FALSE,
          classes = "compact", # compact style for smaller rows
          columnDefs = list(
            list(
              targets = 0,
              render = DT::JS(
                "function(data, type) {
                if (type === 'display') {
                  const date = new Date(data);
                  return date.toLocaleString('en-US', {
                    year: 'numeric',
                    month: 'short',
                    day: 'numeric',
                    hour: '2-digit',
                    minute: '2-digit'
                  });
                }
                return data;
              }"
              )
            )
          )
        ),
        rownames = FALSE,
        # Custom styling
        class = "stripe hover row-border display",
        style = "bootstrap5"
      ) -> dt

      dt <- DT::formatStyle(
        dt,
        "Subject",
        fontWeight = "bold"
      )

      dt <- DT::formatStyle(
        dt,
        0, # Apply to all columns
        target = "row",
        lineHeight = "1.5"
      )

      dt
    })

    # Show the email body when a row is clicked
    output$email_body <- shiny::renderUI({
      # Show welcome message if not authenticated
      if (is.null(values$outlook)) {
        # Different message based on whether API key is set
        if (!is_api_key_set()) {
          shiny::div(
            class = "d-flex flex-column justify-content-center align-items-center h-100",
            style = "color: #0062cc;",
            shiny::div(
              class = "text-center",
              shiny::tags$i(class = "fas fa-lock fa-3x mb-2"),
              shiny::h3("Welcome to ShinyEmailAI", class = "mt-2", style = "font-size: 1.3rem;"),
              shiny::p(
                "First, click the 'Set Anthropic API Key' button to configure your API key",
                class = "mt-2 mb-1",
                style = "font-size: 0.9rem;"
              ),
              shiny::p(
                "Then authenticate with Outlook to get started",
                class = "mt-1",
                style = "font-size: 0.9rem;"
              ),
              shiny::p(
                "This application helps you browse and reply to emails with AI assistance",
                class = "small mt-2",
                style = "font-size: 0.8rem;"
              )
            )
          )
        } else {
          shiny::div(
            class = "d-flex flex-column justify-content-center align-items-center h-100",
            style = "color: #0062cc;",
            shiny::div(
              class = "text-center",
              shiny::tags$i(class = "fas fa-key fa-3x mb-2"),
              shiny::h3("Welcome to ShinyEmailAI", class = "mt-2", style = "font-size: 1.3rem;"),
              shiny::p(
                "Click the 'Authenticate Outlook' button in the top-right corner to get started",
                class = "mt-2",
                style = "font-size: 0.9rem;"
              ),
              shiny::p(
                "This application helps you browse and reply to emails with AI assistance",
                class = "small mt-2",
                style = "font-size: 0.8rem;"
              )
            )
          )
        }
      } else {
        sel <- input$email_table_rows_selected
        if (length(sel) != 1) {
          shiny::div(
            class = "d-flex flex-column justify-content-center align-items-center h-100",
            style = "color: #6c757d;",
            shiny::div(
              class = "text-center",
              shiny::tags$i(class = "fas fa-envelope fa-3x mb-2"),
              shiny::h4("Select an email", class = "mt-2", style = "font-size: 1.1rem;"),
              shiny::p(
                "Choose an email from the inbox to view its content",
                class = "small",
                style = "font-size: 0.8rem;"
              )
            )
          )
        } else {
          email <- values$emails_df[sel, ]
          # Store the selected email ID and body for later use
          email_id <- email$id
          values$selected_email_id <- email_id

          # Store body_with_header for use in new_chat
          values$selected_body_with_header <- email$body_with_header

          # Get the raw email object for attachment handling
          raw_email <- values$emails_raw[[sel]]

          # Store text version of email body for prompt context
          text_email <- tryCatch(
            {
              purrr::map_chr(
                values$emails_raw[sel],
                ~ tidy_email(.x, include_body = "text")$body
              )
            },
            error = function(e) {
              message("Error extracting email text: ", e$message)
              return("")
            }
          )
          values$selected_email_body <- text_email

          # Only retrieve existing chat session if available
          if (!is.null(values$chat_store[[email_id]])) {
            # Update the response textarea with the saved draft for this email
            shiny::updateTextAreaInput(
              session,
              "response_text",
              value = values$chat_store[[email_id]]$draft
            )
          } else {
            # Initialize the structure but don't create the chat object yet
            values$chat_store[[email_id]] <- list(
              chat = NULL, # Chat will be created when "Iterate Draft" is clicked
              draft = "",
              model = NULL
            )
          }

          # Process HTML body to fix image references if needed
          fixed_html_body <- fix_email_images(raw_email)

          # Check for attachments
          has_attachments <- email$has_attachments
          attachment_names <- email$attachment_names

          shiny::tagList(
            shiny::div(
              class = "p-1 border-bottom",
              style = "margin-bottom: 0.3rem;",
              shiny::div(
                class = "d-flex justify-content-between align-items-center",
                shiny::div(
                  shiny::tags$h4(email$subject, class = "mb-1", style = "font-size: 0.95rem; margin-bottom: 0.2rem;"),
                  shiny::div(
                    class = "text-muted small",
                    style = "font-size: 0.75rem;",
                    shiny::div(
                      shiny::tags$span(shiny::tags$strong("From: "), email$from_names)
                    ),
                    shiny::div(
                      shiny::tags$span(
                        shiny::tags$strong("To: "),
                        if (!is.na(email$to_names)) {
                          email$to_names
                        } else {
                          "No recipients"
                        }
                      )
                    ),
                    shiny::div(
                      shiny::tags$span(
                        shiny::tags$strong("CC: "),
                        if (!is.na(email$cc_names)) email$cc_names else "None"
                      )
                    ),
                    shiny::div(
                      shiny::tags$span(
                        shiny::tags$strong("Date: "),
                        tryCatch(
                          {
                            if (is.na(email$sent_datetime)) {
                              "Date unavailable"
                            } else {
                              format(
                                as.POSIXct(email$sent_datetime),
                                "%a, %b %d, %Y, %I:%M %p"
                              )
                            }
                          },
                          error = function(e) {
                            "Date unavailable"
                          }
                        )
                      )
                    ),
                    # Add attachment information if present
                    if (has_attachments && !is.na(attachment_names)) {
                      shiny::div(
                        shiny::tags$span(
                          shiny::tags$strong("Attachments: "),
                          attachment_names
                        )
                      )
                    }
                  )
                ),
                shiny::div(
                  class = "badge bg-secondary",
                  tryCatch(
                    {
                      if (is.na(email$sent_datetime)) {
                        ""
                      } else {
                        format(as.POSIXct(email$sent_datetime), "%a %b %d")
                      }
                    },
                    error = function(e) {
                      ""
                    }
                  )
                )
              )
            ),
            # Email body with fixed images
            shiny::div(
              style = "height: 180px;",
              shiny::tags$iframe(
                srcdoc = fixed_html_body,
                style = "width: 100%; height: 100%; border: none;",
                sandbox = "allow-same-origin allow-scripts" # This might help with some content
              )
            ),
            # Add Archive button
            shiny::div(
              class = "mt-2",
              shiny::actionButton(
                "archive_btn",
                "Mark Read & Archive",
                icon = shiny::icon("archive"),
                class = "btn-outline-secondary"
              )
            ),
            # Attachments section (if any)
            if (has_attachments && !is.na(attachment_names)) {
              shiny::div(
                class = "p-2 mt-2 border-top",
                shiny::h5("Attachments"),
                shiny::div(
                  class = "d-flex flex-wrap gap-2",
                  lapply(
                    email$attachments[[1]],
                    function(att) {
                      shiny::actionButton(
                        inputId = paste0(
                          "download_",
                          gsub("[^a-zA-Z0-9]", "_", att$properties$id)
                        ),
                        label = paste0(
                          att$properties$name,
                          " (",
                          format(round(att$properties$size / 1024), big.mark = ","),
                          " KB)"
                        ),
                        icon = shiny::icon("download"),
                        class = "btn-sm btn-outline-primary"
                      )
                    }
                  )
                )
              )
            }
          )
        }
      }
    })

    # Handle clearing the draft
    shiny::observeEvent(input$clear_draft_btn, {
      if (
        is.null(values$selected_email_id) || is.null(values$selected_email_body)
      ) {
        shiny::showNotification(
          "Please select an email first",
          type = "warning"
        )
        return()
      }

      # Show a progress notification
      id <- shiny::showNotification(
        "Clearing draft and resetting chat...",
        type = "message",
        duration = NULL
      )

      tryCatch(
        {
          # Reset without creating a new chat object
          values$chat_store[[values$selected_email_id]] <- list(
            chat = NULL, # No chat until "Iterate Draft" is clicked
            draft = "",
            model = NULL
          )

          # Clear the response textarea
          shiny::updateTextAreaInput(session, "response_text", value = "")

          # Remove the notification
          shiny::removeNotification(id)

          shiny::showNotification(
            "Draft cleared and chat reset",
            type = "message",
            duration = 3
          )
        },
        error = function(e) {
          shiny::removeNotification(id)
          shiny::showNotification(
            paste("Error clearing draft:", e$message),
            type = "error",
            duration = 10
          )
        }
      )
    })

    # Handle iteration on the draft
    shiny::observeEvent(input$iterate_btn, {
      if (!is_api_key_set()) {
        shiny::showNotification(
          "Please set your Anthropic API key first",
          type = "warning"
        )
        return()
      }

      if (is.null(values$outlook)) {
        shiny::showNotification(
          "Please authenticate with Outlook first",
          type = "warning"
        )
        return()
      }

      if (
        is.null(values$selected_email_id) ||
          is.null(values$chat_store[[values$selected_email_id]])
      ) {
        shiny::showNotification(
          "Please select an email first",
          type = "warning"
        )
        return()
      }

      if (input$user_request == "") {
        shiny::showNotification(
          "Please enter an iteration request",
          type = "warning"
        )
        return()
      }

      # Get the current draft from the textarea
      current_draft <- input$response_text

      # Show a progress notification
      id <- shiny::showNotification(
        "Updating draft...",
        type = "message",
        duration = NULL
      )

      tryCatch(
        {
          # Check if we need to create a chat or change the model
          current_model <- values$chat_store[[values$selected_email_id]]$model
          selected_model <- input$model_choice
          current_chat <- values$chat_store[[values$selected_email_id]]$chat

          # If chat doesn't exist yet or model has changed, create a new chat
          if (
            is.null(current_chat) ||
              is.null(current_model) ||
              current_model != selected_model
          ) {
            # First time creating chat or model change

            # Get turns from the current chat if it exists
            turns <- NULL
            if (!is.null(current_chat)) {
              tryCatch(
                {
                  turns <- current_chat$get_turns(include_system_prompt = FALSE)
                },
                error = function(e) {
                  message(
                    "Could not extract turns from current chat: ",
                    e$message
                  )
                }
              )
            }

            # Use body_with_header from the selected email
            values$chat_store[[values$selected_email_id]]$chat <- new_chat(
              values$selected_body_with_header,
              selected_model,
              turns = turns,
              prompts_dir = prompts_dir,
              email_samples_path = file.path(prompts_dir, "email-samples.md")
            )
            values$chat_store[[values$selected_email_id]]$model <- selected_model

            if (is.null(current_chat)) {
              shiny::showNotification(
                paste("Created new chat with model:", selected_model),
                type = "message",
                duration = 3
              )
            } else {
              shiny::showNotification(
                paste("Switched to model:", selected_model),
                type = "message",
                duration = 3
              )
            }
          }

          # Get the chat for the current email
          current_chat <- values$chat_store[[values$selected_email_id]]$chat

          # Use the iteration prompt from prompts directory
          prompt_path <- file.path(prompts_dir, "prompt-iterate-on-draft.md")
          if (!file.exists(prompt_path)) {
            # Default prompt if file doesn't exist
            prompt_text <- paste(
              "- There is a request.",
              "- If there is a current email draft it is shown here.",
              "- The draft might differ from what you've seen previously.",
              "- Use the draft here going forward.",
              "",
              "[Current draft]",
              "",
              current_draft,
              "",
              "[/Current draft]",
              "",
              "[Request]",
              "",
              input$user_request,
              "",
              "[/Request]",
              sep = "\n"
            )
          } else {
            prompt_text <- ellmer::interpolate_file(
              prompt_path,
              currentDraft = current_draft,
              userRequest = input$user_request
            )
          }

          # Get the updated draft with context if available and there are actual PDF files
          if (!is.null(values$context_files) && length(values$context_files) > 0) {
            # Verify the files exist before trying to use them
            existing_files <- values$context_files[file.exists(values$context_files)]

            if (length(existing_files) > 0) {
              # Create a list of PDF content files
              pdf_contents <- lapply(existing_files, ellmer::content_pdf_file)
              # Add PDF contents to the chat call
              updated_draft <- do.call(
                current_chat$chat,
                c(list(prompt_text), pdf_contents)
              )
            } else {
              # No valid PDF files found, proceed without context
              updated_draft <- current_chat$chat(prompt_text)
            }
          } else {
            updated_draft <- current_chat$chat(prompt_text)
          }

          # Update the textarea with the new draft
          shiny::updateTextAreaInput(session, "response_text", value = updated_draft)

          # Save the updated draft in the chat store
          values$chat_store[[values$selected_email_id]]$draft <- updated_draft

          # Remove the notification
          shiny::removeNotification(id)

          shiny::showNotification(
            "Draft updated successfully",
            type = "message",
            duration = 3
          )
        },
        error = function(e) {
          shiny::removeNotification(id)
          shiny::showNotification(
            paste("Error updating draft:", e$message),
            type = "error",
            duration = 10
          )
        }
      )
    })

    # Save draft when response text changes
    shiny::observeEvent(input$response_text, {
      if (
        !is.null(values$selected_email_id) &&
          !is.null(values$chat_store[[values$selected_email_id]])
      ) {
        # Save the current draft to the chat store
        values$chat_store[[values$selected_email_id]]$draft <- input$response_text
      }
    })

    # Handle creating Outlook "Reply" draft
    shiny::observeEvent(input$create_reply_draft, {
      if (is.null(values$outlook)) {
        shiny::showNotification(
          "Please authenticate with Outlook first",
          type = "warning"
        )
      } else if (is.null(values$selected_email_id)) {
        shiny::showNotification(
          "Please select an email to reply to first",
          type = "warning"
        )
      } else if (input$response_text == "") {
        shiny::showNotification("Please enter text for your response", type = "warning")
      } else {
        # Get the current draft text
        current_draft <- input$response_text

        # Save the current draft to the chat store
        if (!is.null(values$chat_store[[values$selected_email_id]])) {
          values$chat_store[[values$selected_email_id]]$draft <- current_draft
        }

        tryCatch(
          {
            # Show processing notification
            id <- shiny::showNotification(
              "Creating reply draft...",
              type = "message",
              duration = NULL
            )

            # Convert markdown to HTML first
            current_draft_html <- markdown::markdownToHTML(
              text = current_draft,
              fragment.only = TRUE
            )

            # Wrap the HTML in a styled div for email compatibility
            current_draft_html_style <- paste0(
              "<div style='font-size:11pt; font-family:Calibri,sans-serif;'>",
              current_draft_html,
              "</div>"
            )

            # Get the email object and create the reply
            email <- values$outlook$get_inbox()$get_email(values$selected_email_id)
            email$create_reply(current_draft_html_style)
            shiny::showNotification("'Reply' draft created in Outlook", type = "message")
          },
          error = function(e) {
            shiny::removeNotification(id)
            shiny::showNotification(
              paste("Error creating draft:", e$message),
              type = "error"
            )
          }
        )
      }
    })

    # Handle creating Outlook "Reply All" draft
    shiny::observeEvent(input$create_reply_all_draft, {
      if (is.null(values$outlook)) {
        shiny::showNotification(
          "Please authenticate with Outlook first",
          type = "warning"
        )
      } else if (is.null(values$selected_email_id)) {
        shiny::showNotification(
          "Please select an email to reply to first",
          type = "warning"
        )
      } else if (input$response_text == "") {
        shiny::showNotification("Please enter text for your response", type = "warning")
      } else {
        # Get the current draft text
        current_draft <- input$response_text

        # Save the current draft to the chat store
        if (!is.null(values$chat_store[[values$selected_email_id]])) {
          values$chat_store[[values$selected_email_id]]$draft <- current_draft
        }

        tryCatch(
          {
            # Show processing notification
            id <- shiny::showNotification(
              "Creating 'Reply All' draft...",
              type = "message",
              duration = NULL
            )

            # Convert markdown to HTML first
            current_draft_html <- markdown::markdownToHTML(
              text = current_draft,
              fragment.only = TRUE
            )

            # Wrap the HTML in a styled div for email compatibility
            current_draft_html_style <- paste0(
              "<div style='font-size:11pt; font-family:Calibri,sans-serif;'>",
              current_draft_html,
              "</div>"
            )

            # Get the email object and create the reply
            email <- values$outlook$get_inbox()$get_email(values$selected_email_id)
            email$create_reply_all(current_draft_html_style)
            shiny::showNotification("'Reply All' draft created in Outlook", type = "message")
          },
          error = function(e) {
            shiny::removeNotification(id)
            shiny::showNotification(
              paste("Error creating draft:", e$message),
              type = "error"
            )
          }
        )
      }
    })

    shiny::observeEvent(
      input$email_table_rows_selected,
      {
        sel <- input$email_table_rows_selected
        if (length(sel) != 1) {
          return()
        }

        email <- values$emails_df[sel, ]
        if (!email$has_attachments || is.na(email$attachment_names)) {
          return()
        }

        attachments <- email$attachments[[1]]
        if (length(attachments) == 0) {
          return()
        }

        for (att in attachments) {
          local({
            local_att <- att
            local_id <- gsub("[^A-Za-z0-9]", "_", att$properties$id)
            local_btn <- paste0("download_", local_id)

            shiny::observeEvent(
              input[[local_btn]],
              {
                # --------------------------------------------------------------------- #
                # 1. Only handle true file attachments
                # --------------------------------------------------------------------- #
                if (
                  !identical(
                    local_att$properties$`@odata.type`,
                    "#microsoft.graph.fileAttachment"
                  )
                ) {
                  shiny::showNotification(
                    paste(
                      "Attachment",
                      local_att$properties$name,
                      "is a link/item - download not supported."
                    ),
                    type = "warning",
                    duration = 6
                  )
                  return(invisible())
                }

                # --------------------------------------------------------------------- #
                # 2. Build a unique filename using the immutable id
                # --------------------------------------------------------------------- #
                download_dir <- downloads_dir()
                dir.create(download_dir, showWarnings = FALSE, recursive = TRUE)

                safe_name <- local_att$properties$name
                file_path <- file.path(download_dir, safe_name)

                notification_id <- paste0("download_", local_id)
                shiny::showNotification(
                  paste("Downloading", local_att$properties$name, "..."),
                  id = notification_id,
                  type = "message",
                  duration = NULL
                )

                tryCatch(
                  {
                    # ------------------------------------------------------------------- #
                    # 3. Build a unique destination path so existing files are preserved
                    # ------------------------------------------------------------------- #
                    dest_dir <- dirname(file_path)
                    fname <- basename(file_path)
                    base_name <- tools::file_path_sans_ext(fname)
                    ext <- tools::file_ext(fname)

                    # Start with the original name; if it exists, add " (1)", " (2)", etc.
                    candidate <- file.path(dest_dir, fname)
                    counter <- 1
                    while (file.exists(candidate)) {
                      suffix <- sprintf(" (%d)", counter)
                      candidate <- file.path(
                        dest_dir,
                        if (nzchar(ext)) {
                          sprintf("%s%s.%s", base_name, suffix, ext)
                        } else {
                          sprintf("%s%s", base_name, suffix)
                        }
                      )
                      counter <- counter + 1
                    }

                    # Finally download to the unique path
                    local_att$download(dest = candidate, overwrite = FALSE)

                    shiny::removeNotification(notification_id)
                    shiny::showNotification(
                      paste("Downloaded to:", candidate),
                      type = "message",
                      duration = 5
                    )
                  },
                  error = function(e) {
                    shiny::removeNotification(notification_id)
                    shiny::showNotification(
                      paste("Error downloading attachment:", e$message),
                      type = "error",
                      duration = 10
                    )
                  }
                )
              },
              ignoreInit = TRUE,
              ignoreNULL = TRUE
            )
          })
        }
      },
      ignoreInit = TRUE
    )

    # Handle 'Mark Read & Archive' button
    shiny::observeEvent(input$archive_btn,
      {
        shiny::req(values$outlook, values$selected_email_id)
        tryCatch(
          {
            # Retrieve the message from Inbox
            mail <- values$outlook$get_inbox()$get_email(values$selected_email_id)
            # Mark as read
            mail$update(isRead = TRUE)
            # Move to Archive folder using Microsoft Graph API
            # Get the Archive folder ID first
            archive_folder <- values$outlook$get_folder("Archive")

            if (!is.null(archive_folder)) {
              # Move the message to Archive folder using the move() method directly on the mail object
              mail$move(archive_folder)
            } else {
              stop("Archive folder not found")
            }
            shiny::showNotification("Email marked read and moved to Archive.", type = "message")
            # Refresh inbox
            values$emails_raw <- values$outlook$list_emails()
            values$emails_df <- purrr::map_dfr(
              values$emails_raw,
              ~ tidy_email(.x, include_body = "html")
            )
          },
          error = function(e) {
            shiny::showNotification(paste("Error archiving email:", e$message), type = "error")
          }
        )
      },
      ignoreInit = TRUE
    )
  }

  # ---------------------------------------------------------------------------
  # Add JavaScript for dynamic features
  # ---------------------------------------------------------------------------

  js <- "
  $(document).ready(function() {
    // Add hover effects to buttons
    $('.action-button').hover(
      function() { $(this).addClass('shadow-sm'); },
      function() { $(this).removeClass('shadow-sm'); }
    );

    // Set initial font sizes for better readability
    $('#response_text').css('font-size', '13px');
    $('#user_request').css('font-size', '13px');

    // Adjust on window resize
    $(window).resize(function() {
      // Redraw DataTable to adjust column widths
      if ($.fn.dataTable && $.fn.dataTable.tables().length > 0) {
        $($.fn.dataTable.tables()).DataTable().columns.adjust();
      }
    });
  });
  "
  # resizable functionality
  css <- "
  .resizable-card {
    transition: box-shadow 0.3s ease;
    position: relative;
  }
  .resizable-card:hover {
    box-shadow: 0 4px 12px rgba(0,0,0,0.15);
  }
  .resizable-card .card-header {
    cursor: default;
    user-select: none;
  }
  .ui-resizable-handle {
    background-color: rgba(200, 200, 200, 0.3);
    border-radius: 3px;
  }
  .ui-resizable-handle:hover {
    background-color: rgba(0, 123, 255, 0.5);
  }

  /* Email table scrolling styles */
  #inbox_card .dataTables_wrapper {
    height: 100%;
    width: 100%;
    max-width: calc(100% - 2px);
  }

  #inbox_card .dataTables_scrollBody {
    height: calc(100% - 40px) !important;
  }

  #inbox_card table.dataTable {
    margin-top: 0 !important;
    margin-bottom: 0 !important;
  }

  #inbox_card table.dataTable td {
    padding: 8px 8px !important;
    font-size: 0.92rem;
  }

  #inbox_card table.dataTable th {
    padding: 10px 8px !important;
    font-size: 0.92rem;
  }

  /* Style for selected row */
  #inbox_card table.dataTable tbody tr.selected {
    background-color: rgba(0, 98, 204, 0.15) !important;
    font-weight: 500;
  }

  /* Fixed layout styles */
  .right-column {
    display: flex;
    flex-direction: column;
    gap: 1rem;
  }

  .right-column .card {
    flex: 1;
  }
  "

  shiny::shinyApp(
    ui = shiny::tagList(
      shiny::tags$head(
        shiny::tags$link(
          rel = "stylesheet",
          href = "https://code.jquery.com/ui/1.12.1/themes/base/jquery-ui.css"
        ),
        shiny::tags$script(src = "https://code.jquery.com/ui/1.12.1/jquery-ui.min.js"),
        shiny::tags$style(HTML(css)),
        shiny::tags$script(HTML(js))
      ),
      ui
    ),
    server = server,
    options = list(launch.browser = TRUE),
    ...
  )
}
