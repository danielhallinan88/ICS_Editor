

#======================================
#------------------------
Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing

$title = "ICS Editor"
$new_filename = "MODIFIED.ics"
$desktop_path = [Environment]::GetFolderPath("Desktop")
$output_path = $desktop_path + "\" + $new_filename

$form = New-Object System.Windows.Forms.Form;
 $form.Width = 500;
 $form.Height = 700;
 $form.Text = $title;
 $form.StartPosition = [System.Windows.Forms.FormStartPosition]::CenterScreen;

##############Define text label1
 $textLabel1 = New-Object System.Windows.Forms.Label;
 $textLabel1.Left = 25;
 $textLabel1.Top = 15;
 $textLabel1.Width = 50
 $textLabel1.Text = 'ICS File:';


############Define text box1 for input
 $ics_file = New-Object System.Windows.Forms.TextBox;
 $ics_file.Left = 75;
 $ics_file.Top = 10;
 $ics_file.width = 200;
 $ics_file.Text = "";


#############define buttons
 $browse_button = New-Object System.Windows.Forms.Button;
 $browse_button.Left = 280;
 $browse_button.Top = 8;
 $browse_button.Width = 80;
 $browse_button.Text = 'Browse';

 $open_button = New-Object System.Windows.Forms.Button;
 $open_button.Left = 365;
 $open_button.Top = 8;
 $open_button.Width = 40;
 $open_button.Text = 'Open';

#############Listbox
 $event_list = New-Object System.Windows.Forms.ListBox
 $event_list.Left = 25
 $event_list.Top = 40
 $event_list.Width = 435
 $event_list.Height = 500
 $event_list.SelectionMode = 'MultiExtended'

 $search_button = New-Object System.Windows.Forms.Button
 $search_button.Left = 235
 $search_button.Top = 563
 $search_button.Width = 60
 $search_button.Text = 'Search'

 $search_textbox = New-Object System.Windows.Forms.TextBox;
 $search_textbox.Left = 25;
 $search_textbox.Top = 565;
 $search_textbox.width = 200;
 $search_textbox.Text = "";

 $remove_button = New-Object System.Windows.Forms.Button
 $remove_button.Left = 25
 $remove_button.Top = 595
 $remove_button.Width = 125
 $remove_button.Text = 'Remove Selected'

 $save_button = New-Object System.Windows.Forms.Button
 $save_button.Left = 25
 $save_button.Top = 625
 $save_button.Width = 60
 $save_button.Text = 'Save As'

 $text_label2 = New-Object System.Windows.Forms.Label;
 $text_label2.Left = 90
 $text_label2.Top = 630
 $text_label2.Width = 225
 $text_label2.Text = $output_path;



function Create-EventHash {
    param
    (
        [string]$event,
        [int]$id
    )
    $event_hash = @{
        id = 0;
        title = "TITLE GOES HERE";
        raw_text = $event;
        date = "";
    }

    $title_regex = "SUMMARY:([\s\S]+?)\n"
    $results = $event | Select-String $title_regex -AllMatches
    $event_hash.title = $results.Matches.Groups[1]

    $date_regex = "DTSTART[\s\S]+?:(\S+)"
    $results = $event | Select-String $date_regex -AllMatches
    $date = $results.Matches.Groups[1].ToString().Substring(0,8)

    $date_str = $date.Insert(4, '-')
    $date_str = $date_str.Insert(7, '-')

    $event_hash.date = $date_str

    return $event_hash 
}

function Parse-ICS {
    param
    (
        [string]$ics_content
    )

    $lines = $ics_content -split "`r`n"

    $header_output = ""
    foreach($line in $lines){

        if ($line.StartsWith("BEGIN:VEVENT")){
            break
        }
        else {
            $header_output += $line + "`r`n"
        }
    }
    #$header_output.Trim()

    #$header_regex = "BEGIN[\s\S]*END:VTIMEZONE"
    #$header = Select-String -InputObject $ics_contents -AllMatches $header_regex

    #New-Variable -Name header_text -Value $($header.Matches.Groups[0].ToString()) -Scope Script -Force
    New-Variable -Name header_text -Value $($header_output) -Scope Script -Force
    $event_regex = "BEGIN:VEVENT[\s|\S]*?END:VEVENT"
    $events = Select-String -InputObject $ics_contents -AllMatches $event_regex

    New-Variable -Name events_hashes -Value $($events.Matches.Value | foreach { Create-EventHash -event $_ }) -Scope Script -Force
    New-Variable -Name footer -Value "END:VCALENDAR" -Scope Script -Force

    return $header_text, $events_hashes, $footer
}

$browse_button.Add_Click(
    {
        # You can call any function from here to be executed on click.
        $FileBrowser = New-Object System.Windows.Forms.OpenFileDialog -Property @{ InitialDirectory = [Environment]::GetFolderPath('Desktop') }
        $null = $FileBrowser.ShowDialog()
        #Write-Host $FileBrowser.FileName
        $ics_file.Text = $FileBrowser.FileName
        
    }
)

$open_button.Add_Click(
    {
        $ics_path = $ics_file.Text
        Write-Host $ics_path

        [string]$ics_contents = Get-Content -Path $ics_path -Raw

        $header, $events, $footer = Parse-ICS -ics_content $ics_contents

        $titles = $events | foreach { $_.title }
        $titles_sorted = $titles | Sort-Object
        $event_list.Items.AddRange($titles_sorted)

        #$event_list.Sorted = true
        
    }
)

$search_button.Add_Click(
    {
        $search_string = $search_textbox.Text
        write-host $search_string
        $event_list.Items.Contains($search_string)
    }
)

$remove_button.Add_Click(
    {

        $items_to_remove = $event_list.SelectedItems
        $items_to_remove_idx = $event_list.SelectedIndices
        $items_arr = $items_to_remove_idx | foreach { $_ }

        $new_items = @()

        foreach($item in $event_list.Items){
             if (!$items_to_remove.Contains($item)) {
                $new_items += $item
             }
        }
        
        $event_list.Items.Clear()
        $event_list.Items.AddRange($new_items)

    }
)

$save_button.Add_Click(
    {
        write-host $header_text
        Write-Host $events_hashes[0].raw_text
        write-host $footer

        $items_to_write = $event_list.Items
        $items_to_write_hashes = @()

        $final_ics = $header_text
        $final_ics += "`r`n"

        foreach($item in $items_to_write){
            foreach($event in $events_hashes){
                if($item -eq $event.title){
                    $final_ics += $event.raw_text
                    $final_ics += "`r`n"
                    #write-host $event.date, $event.title
                }
            }
        }

        $final_ics += $footer

        Out-File -FilePath $output_path -InputObject $final_ics -Encoding ASCII
        #(gc $output_path) | ? {$_.trim() -ne "" } | set-content $output_path

    }
)



#############Add controls to all the above objects defined

 $form.Controls.Add($browse_button);
 $form.Controls.Add($open_button);
 $form.Controls.Add($textLabel1);
 $form.Controls.Add($ics_file);
 $form.Controls.Add($event_list);

 #$form.Controls.Add($search_button)
 #$form.Controls.Add($search_textbox)
 $form.Controls.Add($save_button)
 $form.Controls.Add($remove_button)
 $form.Controls.Add($text_label2)

 $form.ShowDialog()