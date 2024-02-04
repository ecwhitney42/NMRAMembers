#!/usr/bin/env osascript


on run argv
    if class of argv is list then
        set LocMonth to (item 1 of argv)
    else
        set LocMonth to month of (current date) as string
    end if

    --    set LocMonth to "January"
    -- This script pulls a first name and an email address from Numbers and sends each email from the Apple Mail client one at a time

    -- This is the static text that will be used for all email body text
    set LocContent to "

        I am re-sending the January reports because I fixed the bug (this time for sure) that was causing some of the columns to
        output in floating point format. This is a bug on my end that I thought I had fixed before but it reared it's ugly head.
        My apologies for the hassle. BTW, I am working with NMRA IT on this whole CSV thing since they haven't rolled it out yet
        maybe I can convince them that it really is just a giant pain in the neck and they should reconsider their options... ;-)

        Best Regards,

        -Erich

        Erich Whitney
        NER Interim Office Manager
        NER Eastern Area Director
        HUB Division Director
        I Am The NMRA
      "


    set DistributionFile to "/Users/erich/Desktop/NMRAMembers/config/NMRA_Email_Distribution_List.numbers"
    --set DistributionFile to "/Users/erich/Desktop/NMRAMembers/config/NMRA_Email_Distribution_List_Test.numbers"

    --copy ("Sending NMRA Reports for the month of: " & LocMonth) to stdout

    tell application "Numbers"
        activate
        open the DistributionFile

        -- In Numbers the first table is Table 1.  It is hard to know this because the table most likely does not have its name appear, but the active sheet in a new Number file is Table 1.  There can be multiple tables on a sheet.  Therefore you need to designate the document (file) sheet, and table before you can access attributes
        tell table 1 of active sheet of front document

            -- For each row for the count of cells in the first column
            repeat with i from 1 to count of cells of column 1

                -- Get the values of the Name and Address for each row and store locally
                set LocID to value of cell i of column "A"
                set LocRegion to value of cell i of column "B"
                set LocDivision to value of cell i of column "C"
                set LocLName to value of cell i of column "D"
                set LocFName to value of cell i of column "E"
                set LocAddress to value of cell i of column "F"
                set LocCategory to value of cell i of column "G"
                set LocFile to value of cell i of column "H"

                set LocPosixFile to POSIX path of LocFile
                set LocFolder to (("/Users/erich/Desktop/NMRAMembers/release/") & LocMonth)
                set LocAttachment to (POSIX path of (LocFolder as text)) & LocPosixFile
--                set LocAttachment to POSIX file ("/Users/erich/Desktop/NMRAMembers/release/January/NER_Region.zip")
                -- don't send the mail if the attachment doesn't exist
                delay 1
                tell application "Finder"
                    delay 1
                    if exists LocAttachment as POSIX file then

                        tell application "Mail"

                            -- Create new email message
                            set MyMessage to make new outgoing message

                            -- Set the attributes of the email
                            set subject of MyMessage to "Monthly NMRA Roster Reports"
                            -- Append the first name from Numbers to the salutation
                            set content of MyMessage to "Dear " & LocFName & "," & LocContent
                            set sender of MyMessage to "ecwhitney@icloud.com"

                            tell MyMessage
                                -- New email message to each recipient email address
                                make new to recipient at end of to recipients with properties {address:LocAddress}
                                -- Attache the image file as the signature
                                make new attachment with properties {file name:LocAttachment}
                            end tell

                            -- Make sure there is time to read the image attachment and attach it before sending
                            delay 1

                            -- All information is set, send the message
                           send MyMessage

                            -- Make sure the mail client is able to beffer by delaying until the next send
                            delay 10

                        end tell
                    else
                        delay 1
                        copy "Attachment " & LocAttachment & " doesn't exist, not sending email to " & LocFName & " " & LocLName to stdout
                        delay 1
                    end if
                end tell
            end repeat

        end tell

    end tell
end run
