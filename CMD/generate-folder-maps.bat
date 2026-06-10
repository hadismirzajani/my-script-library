:: ==============================================================================
:: SYNOPSIS: Comprehensive Folder Directory and Structure Exporter.
:: DESCRIPTION: Navigates to a target path, generates a raw file list, 
::              creates a visual tree map, and automatically opens both results.
:: ==============================================================================

# Step 1: Navigate to the target directory
# Replace the path below with the folder you want to scan
cd /d "C:\Your\Folder\Path"

# Step 2: Generate a flat, recursive list of all file paths
dir /b /s > file.txt

# Step 3: Open the flat file list immediately for viewing
start file.txt

# Step 4: Create an alphabetical, visual diagram of the directory tree
tree /f /a > tree_list.txt

# Step 5: Open the visual tree map immediately for viewing
start tree_list.txt

# ==============================================================================
# FULL EXPLANATION OF EACH COMMAND PART
# ==============================================================================
#
# 1. cd /d "path"
#    - cd : Changes your current working folder location to a new one.
#    - /d : Swaps both the drive letter and the folder path at the same time.
#    - "path" : Uses quotation marks so paths with spaces do not cause errors.
#
# 2. dir /b /s > file.txt
#    - dir : The core command that reads and displays folder contents.
#    - /b : "Bare format". Strips out file sizes, dates, and header text.
#    - /s : "Subdirectories". Looks recursively inside every single subfolder.
#    - > : Redirection arrow. Saves the text to a file instead of the screen.
#    - file.txt : The text document where your raw list is saved.
#
# 3. start file.txt & start tree_list.txt
#    - start : Tells Windows to launch a file using its default system program.
#    - file.txt / tree_list.txt : Opens these text files instantly in Notepad.
#
# 4. tree /f /a > tree_list.txt
#    - tree : A built-in tool that builds a visual map of folders using lines.
#    - /f : Forces the map to show individual file names, not just folder names.
#    - /a : "ASCII mode". Draws the map using basic text symbols (+---, |, \---).
#           This ensures the layout looks straight and neat on GitHub and Notepad.
#    - > : Saves the visual map layout straight into your text file.
#    - tree_list.txt : The text document where your visual diagram is saved.
