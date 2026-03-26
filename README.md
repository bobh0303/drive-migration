# Google Drive Migration tools

This repo has tools and data used in the migration of WSTech's shared Google docs to a Shared Drive

## Step 1

GTIS Google Admins used a tool that walked the WSTech Team shared folder and created a spreadsheet of all the files and folders it could see.

Originally run as the WSTech Director, we discovered there were files and folders that the director's account couldn't see.

We had GTIS re-run the inventory tool again, once as Bob H, once as Victor G, and once as Peter M, giving us 4 "views" of the shared folder tree.

The four resulting spreadsheets were downloaded as CSV files named `inventory*.csv`

## Step 2

The four "views" were merged using a custom python script, `merge-em.py`, resulting in a `master.csv` file.

## Step 3

`master.csv` was processed by `inventory.py` to create a more usefully formated Excel spreadsheet. Some of the processing included:
- combining item name and webviewlink fields into a single hyperlink
- Providing hyperlinks for the containing folders
- Simplifying permissions into lists of owners, writers, commenters, and viewers
- within those lists, replace our team members' emails with initials, making it more compact but also making external collaborators more obvious
- Comparing owner, writer, commenter, and viewer lists with what we expect for our team, thus identifying potentially confidential documents, and identifying all the distinct combinations of access rights in case that would be helpful.

The result was output to Google as `output CURRENT.xslx` and has been used extensivly in analyzing how to organize the moving of documents to the new Shared Drive, and keeping track of progress.

## Step 4

We requested from GTIS a new inventory, this time of the WSTech Team Shared Drive. The goal is to compare the "before" and "after" pictures, looking particularly for:
- missing documents
- any sharing _outside_ of our team that might have not propogate properly

We've received the inventory and found it's format different enough that we have to rewrite some code.

