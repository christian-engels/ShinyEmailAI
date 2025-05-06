#' Create context structure
#'
#' Creates the necessary directory structure for context and prompts folders
#' at a user-specified location. This is needed for the application to properly
#' utilize context when drafting emails.
#'
#' @param base_dir The base directory where the folders should be created
#' @param sample_contexts Whether to create sample context subdirectories
#' @param copy_prompts Whether to copy the default prompts to the prompts folder
#' @return Invisibly returns the paths of created directories
#' @export
#' @examples
#' \dontrun{
#' # Create context and prompts folders in your home directory
#' email_create_context("~/ShinyEmailAI_Data")
#'
#' # Create with sample contexts and prompts
#' email_create_context("~/ShinyEmailAI_Data", sample_contexts = TRUE, copy_prompts = TRUE)
#' }
email_create_context <- function(base_dir, sample_contexts = TRUE, copy_prompts = TRUE) {
  # Expand path if necessary (e.g., for ~ tilde expansion)
  base_dir <- normalizePath(base_dir, mustWork = FALSE)

  # Create the base directory if it doesn't exist
  if (!dir.exists(base_dir)) {
    dir.create(base_dir, recursive = TRUE)
    message("Created base directory: ", base_dir)
  }

  # Create context directory
  context_dir <- file.path(base_dir, "context")
  if (!dir.exists(context_dir)) {
    dir.create(context_dir)
    message("Created context directory: ", context_dir)
  }

  # Create prompts directory
  prompts_dir <- file.path(base_dir, "prompts")
  if (!dir.exists(prompts_dir)) {
    dir.create(prompts_dir)
    message("Created prompts directory: ", prompts_dir)
  }

  # Create sample context subdirectories if requested
  if (sample_contexts) {
    sample_folders <- c(
      "context-folder-1",
      "context-folder-2",
      "no-context"
    )

    for (folder in sample_folders) {
      folder_path <- file.path(context_dir, folder)
      if (!dir.exists(folder_path)) {
        dir.create(folder_path)
        message("Created sample context folder: ", folder_path)
      }
    }

    # Create a placeholder file in no-context folder
    no_context_file <- file.path(context_dir, "no-context", "do-not-change-this-folder.md")
    if (!file.exists(no_context_file)) {
      writeLines(
        "# Do Not Change This Folder\n\nThis folder is used by ShinyEmailAI as a placeholder for 'no context' operation. Please do not add files to this folder.",
        con = no_context_file
      )
    }
  }

  # Copy prompt templates if requested
  if (copy_prompts) {
    # Default prompts that should be included
    prompt_templates <- c(
      "prompt-system.md",
      "prompt-iterate-on-draft.md",
      "email-samples.md"
    )

    # Get the path to installed package prompt templates
    pkg_prompt_dir <- system.file("extdata", "prompts", package = "ShinyEmailAI")

    if (nchar(pkg_prompt_dir) == 0) {
      warning("Could not find prompt templates in package. Default prompts will not be copied.")
    } else {
      for (template in prompt_templates) {
        src_file <- file.path(pkg_prompt_dir, template)
        dst_file <- file.path(prompts_dir, template)

        if (file.exists(src_file) && !file.exists(dst_file)) {
          file.copy(src_file, dst_file)
          message("Copied prompt template: ", template)
        }
      }
    }
  }

  # Return paths invisibly
  invisible(list(
    base_dir = base_dir,
    context_dir = context_dir,
    prompts_dir = prompts_dir
  ))
}

#' @noRd
get_context_prompts_dirs <- function(base_dir = NULL) {
  if (is.null(base_dir)) {
    base_dir <- getwd()
  }

  context_dir <- file.path(base_dir, "context")
  prompts_dir <- file.path(base_dir, "prompts")

  # Check if directories exist
  if (!dir.exists(context_dir)) {
    warning("Context directory not found: ", context_dir)
  }

  if (!dir.exists(prompts_dir)) {
    warning("Prompts directory not found: ", prompts_dir)
  }

  list(
    context_dir = context_dir,
    prompts_dir = prompts_dir
  )
}

#' @noRd
find_context_files <- function(context_name, base_dir = NULL) {
  dirs <- get_context_prompts_dirs(base_dir)
  context_dir <- file.path(dirs$context_dir, context_name)

  if (!dir.exists(context_dir)) {
    warning("Context directory not found: ", context_dir)
    return(character(0))
  }

  # Find PDF files in the context directory
  pdf_files <- list.files(context_dir, pattern = "\\.pdf$", full.names = TRUE)

  return(pdf_files)
}

#' @noRd
list_context_folders <- function(base_dir = NULL) {
  dirs <- get_context_prompts_dirs(base_dir)

  if (!dir.exists(dirs$context_dir)) {
    warning("Context directory not found: ", dirs$context_dir)
    return(character(0))
  }

  # List subdirectories of context directory
  folders <- list.dirs(dirs$context_dir, full.names = FALSE, recursive = FALSE)

  return(folders)
}
