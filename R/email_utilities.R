# email_utilities.R: Functions for handling Outlook email attachments and images

#' @noRd
save_email_attachments <- function(email, dir = tempdir()) {
  # Validate email object
  if (!inherits(email, "ms_outlook_email")) {
    stop("`email` must be an `ms_outlook_email` object", call. = FALSE)
  }

  # Create the directory if it doesn't exist
  if (!dir.exists(dir)) {
    dir.create(dir, recursive = TRUE)
  }

  # Get attachments
  attachments <- email$list_attachments(n = Inf)

  if (length(attachments) == 0) {
    message("No attachments found")
    return(invisible(character(0)))
  }

  # Download each attachment
  saved_files <- character(length(attachments))

  for (i in seq_along(attachments)) {
    att <- attachments[[i]]

    # Get file info
    file_name <- att$properties$name
    file_path <- file.path(dir, file_name)

    # Download attachment
    content <- att$get_content()
    writeBin(content, file_path)

    saved_files[i] <- file_path
    message("Saved attachment: ", file_path)
  }

  return(saved_files)
}

#' @noRd
get_attachment_bytes <- function(att) {
  props <- att$properties

  ## 1) If Graph already returned contentBytes, decode and return ----
  if (!is.null(props$contentBytes) && nzchar(props$contentBytes)) {
    return(openssl::base64_decode(gsub("[\r\n]", "", props$contentBytes)))
  }

  ## 2) Otherwise hit the /$value endpoint ---------------------------
  # this returns an httr response; pull the body back as raw
  res <- att$do_operation("$value", http_status_handler = "pass")
  if (httr::status_code(res) < 300) {
    return(httr::content(res, as = "raw"))
  }

  # nothing worked → return NULL so caller can skip this attachment
  warning(sprintf("Failed to get attachment bytes via /$value endpoint for attachment ID %s. Status: %d", props$id, httr::status_code(res)))
  NULL
}

#' @noRd
fix_email_images <- function(email) {
  # Ensure email is the correct type
  if (!inherits(email, "ms_outlook_email")) {
    stop("`email` must be an `ms_outlook_email` object", call. = FALSE)
  }

  # Get HTML content directly from the email properties
  html <- email$properties$body$content
  content_type <- email$properties$body$contentType

  # Only process HTML content
  if (is.null(html) || !identical(tolower(content_type), "html")) {
    return(html %||% "")
  }

  if (!grepl("cid:", html, fixed = TRUE)) {
    # message("No CID references found in email body.")
    return(html)
  }

  # message("Email contains CID references, attempting to fix...")

  # List all attachments
  attachments <- tryCatch(email$list_attachments(n = Inf), error = function(e) {
    warning("Could not list attachments: ", e$message)
    list()
  })

  if (length(attachments) == 0) {
    message("No attachments found.")
    return(html)
  }

  # For each inline attachment, replace references in the HTML
  for (att in attachments) {
    props <- att$properties

    # Only process inline file attachments
    if (!isTRUE(props$isInline) || !identical(props$`@odata.type`, "#microsoft.graph.fileAttachment")) {
      next
    }

    # Get content ID, properly formatted
    content_id <- props$contentId
    if (!is.null(content_id)) {
      # Strip angle brackets if present
      content_id <- gsub("^<|>$", "", content_id)

      # Get the attachment bytes
      bytes <- tryCatch(get_attachment_bytes(att), error = function(e) {
        warning("Error getting attachment bytes: ", e$message)
        NULL
      })

      if (!is.null(bytes)) {
        # Create data URI
        b64 <- openssl::base64_encode(bytes, linebreaks = FALSE)
        mime <- props$contentType %||% mime::guess_type(props$name) %||% "application/octet-stream"
        data_uri <- paste0("data:", mime, ";base64,", b64)

        # Try multiple patterns for CID references
        patterns <- c(
          # Standard img tag with quotes
          sprintf('<img[^>]*src="cid:%s"[^>]*>', content_id),
          sprintf("<img[^>]*src='cid:%s'[^>]*>", content_id),
          # Any attribute with this CID
          sprintf("cid:%s", content_id)
        )

        for (pattern in patterns) {
          if (grepl(pattern, html, fixed = (pattern == sprintf("cid:%s", content_id)))) {
            if (pattern == sprintf("cid:%s", content_id)) {
              # Simple string replacement
              html <- gsub(pattern, data_uri, html, fixed = TRUE)
              # message("Fixed CID reference: ", content_id, " (simple replacement)")
            } else {
              # Replace just the src attribute in the img tag
              matches <- gregexpr(pattern, html, perl = TRUE)
              if (matches[[1]][1] != -1) {
                match_strings <- regmatches(html, matches)[[1]]
                for (match in match_strings) {
                  new_tag <- gsub(
                    sprintf('src=["\']cid:%s["\']', content_id),
                    sprintf('src="%s"', data_uri),
                    match,
                    perl = TRUE
                  )
                  html <- gsub(match, new_tag, html, fixed = TRUE)
                  # message("Fixed CID reference: ", content_id, " (attribute replacement)")
                }
              }
            }
          }
        }
      }
    }
  }

  return(html)
}


#' @noRd
display_email_with_attachments <- function(email, download_dir = tempdir()) {
  # --- Input Validation and Object Preparation ---
  raw_email <- NULL
  email_has_attachments <- FALSE
  attachment_names <- NA_character_
  email_body_raw <- ""
  email_attachments <- list()
  email_subject <- "No Subject"
  email_from <- list(name = "Unknown Sender", address = "")
  email_to <- list()
  email_sent_time <- NA

  if (inherits(email, "data.frame")) {
    # Assuming the dataframe row contains a column 'raw_email' with the ms_outlook_email object
    if ("raw_email" %in% names(email) && inherits(email$raw_email[[1]], "ms_outlook_email")) {
      raw_email <- email$raw_email[[1]]
    } else {
      stop("Provided dataframe row must contain a 'raw_email' column with an 'ms_outlook_email' object.")
    }
    # Extract other info preferably from the raw object for consistency
    email_has_attachments <- raw_email$properties$hasAttachments %||% FALSE
    attachment_list <- if (email_has_attachments) raw_email$list_attachments(n = Inf) else list()
    email_attachments <- attachment_list # Keep the list of attachment objects
    attachment_names <- if (length(attachment_list)) {
      paste(vapply(attachment_list, function(a) a$properties$name %||% "Unnamed", ""), collapse = "; ")
    } else {
      NA_character_
    }
    email_body_raw <- raw_email$properties$body$content %||% ""
    email_subject <- raw_email$properties$subject %||% "No Subject"
    email_from <- raw_email$properties$from$emailAddress %||% list(name = "Unknown Sender", address = "")
    email_to <- raw_email$properties$toRecipients %||% list()
    email_sent_time <- raw_email$properties$sentDateTime
  } else if (inherits(email, "ms_outlook_email")) {
    raw_email <- email
    email_has_attachments <- raw_email$properties$hasAttachments %||% FALSE
    attachment_list <- if (email_has_attachments) {
      tryCatch(raw_email$list_attachments(n = Inf), error = function(e) {
        warning("Could not list attachments: ", e$message)
        list()
      })
    } else {
      list()
    }
    email_attachments <- attachment_list # Keep the list of attachment objects
    attachment_names <- if (length(attachment_list)) {
      paste(vapply(attachment_list, function(a) a$properties$name %||% "Unnamed", ""), collapse = "; ")
    } else {
      NA_character_
    }
    email_body_raw <- raw_email$properties$body$content %||% ""
    email_subject <- raw_email$properties$subject %||% "No Subject"
    email_from <- raw_email$properties$from$emailAddress %||% list(name = "Unknown Sender", address = "")
    email_to <- raw_email$properties$toRecipients %||% list()
    email_sent_time <- raw_email$properties$sentDateTime
  } else {
    stop("`email` must be either an `ms_outlook_email` object or a dataframe row containing one.")
  }

  # --- Process HTML Body for Inline Images ---
  # Ensure we have the raw email object to pass to fix_email_images
  if (!is.null(raw_email)) {
    # Only attempt to fix images if the body is HTML
    if (identical(tolower(raw_email$properties$body$contentType %||% ""), "html")) {
      message("Processing email body for inline images...")
      email_body_processed <- tryCatch(
        fix_email_images(raw_email),
        error = function(e) {
          warning("Error processing inline images: ", e$message)
          email_body_raw # Fallback to raw body on error
        }
      )
    } else {
      message("Email body is not HTML, skipping image processing.")
      # If plain text, ensure it's properly escaped for HTML display
      email_body_processed <- htmltools::htmlEscape(email_body_raw)
      # Optionally wrap in <pre> for better formatting
      email_body_processed <- htmltools::tags$pre(email_body_processed)
    }
  } else {
    # Should not happen if validation above is correct, but as a fallback:
    email_body_processed <- htmltools::htmlEscape(email_body_raw)
    email_body_processed <- htmltools::tags$pre(email_body_processed)
  }


  # --- Format Date ---
  formatted_date <- "Date Unknown"
  if (!is.null(email_sent_time) && !is.na(email_sent_time)) {
    parsed_time <- tryCatch(as.POSIXct(email_sent_time, tz = "UTC"), error = function(e) NULL) # Assuming UTC from Graph API
    if (!is.null(parsed_time)) {
      formatted_date <- format(parsed_time, "%a, %b %d, %Y, %I:%M %p %Z")
    }
  }

  # --- Format Recipients ---
  format_recipient <- function(rec) {
    name <- rec$emailAddress$name %||% ""
    addr <- rec$emailAddress$address %||% ""
    if (nzchar(name) && nzchar(addr)) {
      paste0(name, " <", addr, ">")
    } else {
      addr # Fallback to address if name is missing
    }
  }
  to_recipients_formatted <- paste(vapply(email_to, format_recipient, ""), collapse = "; ")


  # --- Create UI Elements ---
  email_display <- htmltools::tagList(
    # Email header information
    htmltools::div(
      class = "email-header",
      style = "border-bottom: 1px solid #ccc; margin-bottom: 15px; padding-bottom: 10px;",
      htmltools::h3(email_subject),
      htmltools::p(
        htmltools::strong("From: "), htmltools::htmlEscape(email_from$name), " <", htmltools::htmlEscape(email_from$address), ">",
        htmltools::br(),
        htmltools::strong("To: "), htmltools::htmlEscape(to_recipients_formatted),
        htmltools::br(),
        htmltools::strong("Date: "), formatted_date
      )
    ),

    # Email body with fixed images (or escaped text)
    htmltools::div(
      class = "email-body",
      # Use htmltools::HTML only if we processed HTML, otherwise use the <pre> tag or escaped text directly
      if (identical(tolower(raw_email$properties$body$contentType %||% ""), "html")) {
        htmltools::HTML(email_body_processed)
      } else {
        email_body_processed # This is already escaped/tagged if plain text
      }
    ),

    # Attachments section (if any)
    if (email_has_attachments && length(email_attachments) > 0) {
      htmltools::div(
        class = "email-attachments",
        style = "margin-top: 20px; padding-top: 10px; border-top: 1px solid #eee;",
        htmltools::hr(),
        htmltools::h4("Attachments"),
        htmltools::tags$ul(
          lapply(
            email_attachments,
            function(att) {
              att_props <- att$properties
              att_name <- att_props$name %||% "Unnamed Attachment"
              att_id <- att_props$id %||% ""
              att_size_kb <- tryCatch(round(att_props$size / 1024), error = function(e) NA)
              att_size_formatted <- if (!is.na(att_size_kb)) {
                paste0(" (", format(att_size_kb, big.mark = ","), " KB)")
              } else {
                ""
              }

              # Basic check if ID seems valid for download button
              can_download <- nzchar(att_id)

              htmltools::tags$li(
                htmltools::htmlEscape(att_name),
                att_size_formatted,
                if (can_download) {
                  # Ensure name is properly escaped for JavaScript string
                  js_name <- gsub("'", "\\'", att_name, fixed = TRUE)
                  js_name <- gsub("\"", "\\\"", js_name, fixed = TRUE)
                  htmltools::tags$button(
                    "Download",
                    # Pass necessary info to Shiny input
                    onclick = sprintf(
                      "Shiny.setInputValue('download_attachment', { id: '%s', name: '%s', nonce: Math.random() }, { priority: 'event' });",
                      att_id,
                      js_name
                    ),
                    class = "btn btn-sm btn-primary ms-2", # Using bootstrap classes assuming Shiny uses it
                    style = "margin-left: 10px;"
                  )
                } else {
                  htmltools::tags$span(" (Download unavailable)", style = "color: grey; margin-left: 10px;")
                }
              )
            }
          ) # end lapply
        ) # end ul
      ) # end div
    } # end if attachments
  ) # end tagList

  return(email_display)
}

# Add `%||%` helper if not already available
`%||%` <- function(x, y) if (is.null(x) || length(x) == 0) y else x
