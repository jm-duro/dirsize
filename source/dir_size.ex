without warning
include std/text.e
include std/search.e
include std/filesys.e
include std/error.e
include std/os.e
include std/convert.e
include std/sequence.e
include common.e
include win32lib_r2.ew
include xlsxwriter.e

constant
  resourceDir = InitialDir & "icons"

constant
  APPL = "Dir Size",
  VERS = "9.1",
  YEAR = "2023"

sequence ov

integer f_out, f_probs
object VOID
atom
  LV, TV, problems, lbl, iExw, ptr, mask, closefolder, openfolder

integer
  Win, SB, btnExpand, btnStart, tvs, by_path

sequence
  folders, location, list, tooLong, toZip

folders = {}

atom
  G20  = 20971520000,
  G10  = 10485760000,
  G5   = 5242880000,
  G2   = 2097152000,
  G1   = 1048576000,
  M500 = 524288000,
  M200 = 209715200,
  M100 = 104857600

------------------------------------------------------------------------------

function crash(object x)
  log_puts("-- Crash ---------------------------------------------------------------------\n")
  analyze_object(folders, "folders", f_debug)
  close(f_debug)
  return 0
end function
crash_routine(routine_id("crash"))

------------------------------------------------------------------------------

function formatNb(atom bytes) -- Returns formatted size
  sequence s
  atom     a

  s = {"KB","MB","GB","TB"}
  for i = 4 to 1 by -1 do
    a = power(1024,i)
    if bytes >= a then
      return sprintf("%2.f %s",{bytes/a, s[i]})
    end if
  end for
  return sprintf("%d bytes", bytes)
end function

------------------------------------------------------------------------------

--converts an ASCII character to its unicode equivalent
function unicode(integer c)
  sequence bits, u8

    if c >= 65536 then     -- 11110xxx 10xxxxxx 10xxxxxx 10xxxxxx
      bits = reverse(int_to_bits(c,21))
      u8 = {{1,1,1,1,0}&bits[1..3],{1,0}&bits[4..9],{1,0}&bits[10..15],{1,0}&bits[16..21]}
      return bits_to_int(reverse(u8[1]))&bits_to_int(reverse(u8[2]))&bits_to_int(reverse(u8[3]))&bits_to_int(reverse(u8[4]))
   elsif c >= 2048 then  -- 1110xxxx 10xxxxxx 10xxxxxx
      bits = reverse(int_to_bits(c,16))
      u8 = {{1,1,1,0}&bits[1..4],{1,0}&bits[5..10],{1,0}&bits[11..16]}
      return bits_to_int(reverse(u8[1]))&bits_to_int(reverse(u8[2]))&bits_to_int(reverse(u8[3]))
    elsif c >= 128 then   -- 110xxxxx 10xxxxxx
      bits = reverse(int_to_bits(c,11))
      u8 = {{1,1,0}&bits[1..5],{1,0}&bits[6..11]}
      return bits_to_int(reverse(u8[1]))&bits_to_int(reverse(u8[2]))
    else                -- 0xxxxxxx
      return {c}
    end if
end function

------------------------------------------------------------------------------

--converts an ASCII string to an UTF-8 string
function ascii_to_utf8(sequence s)
  sequence res = ""
  for i = 1 to length(s) do
    res &= unicode(s[i])
  end for
  return res
end function
------------------------------------------------------------------------------

function dirsize(sequence parent, sequence rootdir)
-- Returns sum of all file sizes from rootdir and all subdirs.
  object   x
  sequence s, res
  atom     bytes
  integer  nb_files, recent, lg, n

  -- puts(f_debug, "parent = " & parent & ", rootdir = " & rootdir & "\n")
  folders = append(folders, {ascii_to_utf8(parent), ascii_to_utf8(rootdir), 0, 0, 0})
  n = length(folders)
  setText({SB, 2}, rootdir)
  recent = 0
  bytes = 0
  nb_files = 0
  lg = length(rootdir)
  if lg > 250 then
    tooLong = append(tooLong, {rootdir, lg})
    toZip = append_if_new(toZip, parent)
    return {bytes, nb_files, recent}
  end if
  x = dir(rootdir & SLASH & "*.*")        -- Get DIR information of all files
  if sequence(x) then             -- Exclude forbidden DIRs
    for i = 1 to length(x) do   -- Loop through DIR content
      s = x[i]
--      analyze_object(s[D_NAME], "s[D_NAME]", f_debug)
      if find('d',s[D_ATTRIBUTES]) then  -- Check for SubDIRs
        if not (equal(s[D_NAME],".") or equal(s[D_NAME],"..") or begins("~",s[D_NAME])) then
          res = dirsize(rootdir, rootdir & SLASH & s[D_NAME])
          bytes += res[1]
          nb_files += res[2]
          if (res[3] > recent) then recent = res[3] end if
        end if
      else
        lg = length(rootdir & SLASH & s[D_NAME])
        if lg > 250 then
          tooLong = append(tooLong, {rootdir & SLASH & s[D_NAME], lg})
          toZip = append_if_new(toZip, rootdir)
        end if
        bytes += s[D_SIZE]       -- Sum Bytes!
        nb_files += 1       -- inc nb_files!
        if not equal(lower(s[D_NAME]),"thumbs.db") and
           (s[D_YEAR] > recent) then
          recent = s[D_YEAR]
        end if
      end if
    end for
  end if
  -- analyze_object(folders, "folders", f_debug)
  folders[n][3] = bytes
  folders[n][4] = nb_files
  folders[n][5] = recent
  if get_key() = 27 then          -- Choice to abort
    setText({SB, 2},"Aborted by ESC key. Press any key . . .")
    sleep(2)
    abort(0)
  end if
  return {bytes, nb_files, recent}
end function

------------------------------------------------------------------------------

procedure Click_btnStart(integer self, integer event, sequence parms)
  sequence rootdir, res, s, d, folderName, outFile
  atom t0
  integer l, n, lg, folderYear
  atom workbook, resultSheet, problemsSheet

  rootdir = selectDirectory("Select directory",BIF_USENEWUI,0,0) & SLASH
  if length(rootdir)>0 then
    d = date()
    s = sprintf("%d-%02d-%02d_%02d-%02d-%02d",
                {d[1]+1900,d[2],d[3],d[4],d[5],d[6]})
    outFile = InitialDir & "analysis_" & s & ".xlsx"
    workbook = new_workbook(outFile)
    resultSheet = workbook_add_worksheet(workbook, "Results")
    worksheet_set_column(resultSheet, {0, 0}, 180, NULL, NULL)
    worksheet_set_column(resultSheet, {1, 4}, 12, NULL, NULL)
    atom formatBold = workbook_add_format(workbook)
    format_set_bold    (formatBold)
    atom over20G    = workbook_add_format(workbook)
    format_set_pattern (over20G, LXW_PATTERN_SOLID)
    format_set_bg_color(over20G, 0xFF0000)
    atom over10G    = workbook_add_format(workbook)
    format_set_pattern (over10G, LXW_PATTERN_SOLID)
    format_set_bg_color(over10G, 0xFF2020)
    atom over5G     = workbook_add_format(workbook)
    format_set_pattern (over5G, LXW_PATTERN_SOLID)
    format_set_bg_color(over5G, 0xFF4040)
    atom over2G     = workbook_add_format(workbook)
    format_set_pattern (over2G, LXW_PATTERN_SOLID)
    format_set_bg_color(over2G, 0xFF6060)
    atom over1G     = workbook_add_format(workbook)
    format_set_pattern (over1G, LXW_PATTERN_SOLID)
    format_set_bg_color(over1G, 0xFF8080)
    atom over500M   = workbook_add_format(workbook)
    format_set_pattern (over500M, LXW_PATTERN_SOLID)
    format_set_bg_color(over500M, 0xFFA0A0)
    atom over200M   = workbook_add_format(workbook)
    format_set_pattern (over200M, LXW_PATTERN_SOLID)
    format_set_bg_color(over200M, 0xFFC0C0)
    atom over100M   = workbook_add_format(workbook)
    format_set_pattern (over100M, LXW_PATTERN_SOLID)
    format_set_bg_color(over100M, 0xFFE0E0)
    worksheet_write_string( resultSheet, {0, 0}, "Name", formatBold )
    worksheet_write_string( resultSheet, {0, 1}, "Bytes", formatBold )
    worksheet_write_string( resultSheet, {0, 2}, "Size", formatBold )
    worksheet_write_string( resultSheet, {0, 3}, "Number", formatBold )
    worksheet_write_string( resultSheet, {0, 4}, "Most recent", formatBold )
    problemsSheet = workbook_add_worksheet(workbook, "Problems")
    worksheet_set_column(problemsSheet, {0, 0}, 180, NULL, NULL)

    t0 = time()
    setText({SB, 1},"ESC to stop.")
--    puts(f_debug, "rootdir = " & rootdir & "\n")
    rootdir = driveid(rootdir) & ":" & dirname(rootdir)
--    puts(f_debug, "rootdir = " & rootdir & "\n")
    tooLong = {}
    toZip = {}
    folders = append(folders, {"", "", 0, 0, 0})
    res = dirsize("", rootdir)
--    analyze_object(folders, "folders1", f_debug)
    lg = length(folders)
    l = 0
    while l < lg do
      l += 1
      folderName = dirname(folders[l][2])
      if length(folderName) = 0 then continue end if
      folderYear = folders[l][5]
      n = l+1
      while n <= lg do
        if begins(folderName, dirname(folders[n][2])) and (folders[n][5] = folderYear) then
          folders = remove(folders, n, n)
          lg = length(folders)
        else
          exit
        end if
      end while
    end while
--    analyze_object(folders, "folders2", f_debug)
    integer year = d[1]+1900
    for i = 2 to length(folders) do
--      puts(f_debug, unicode_to_ascii(folders[i][2]) & "\n")
--      worksheet_write_string( resultSheet, {i-1, 0}, unicode_to_ascii(folders[i][2]), NULL)
      worksheet_write_string( resultSheet, {i-1, 0}, folders[i][2], NULL)
      if folders[i][3] >= G20 then
        worksheet_write_number(resultSheet, {i-1, 1}, folders[i][3], over20G)
      elsif folders[i][3] >= G10 then
        worksheet_write_number(resultSheet, {i-1, 1}, folders[i][3], over10G)
      elsif folders[i][3] >= G5 then
        worksheet_write_number(resultSheet, {i-1, 1}, folders[i][3], over5G)
      elsif folders[i][3] >= G2 then
        worksheet_write_number(resultSheet, {i-1, 1}, folders[i][3], over2G)
      elsif folders[i][3] >= G1 then
        worksheet_write_number(resultSheet, {i-1, 1}, folders[i][3], over1G)
      elsif folders[i][3] >= M500 then
        worksheet_write_number(resultSheet, {i-1, 1}, folders[i][3], over500M)
      elsif folders[i][3] >= M200 then
        worksheet_write_number(resultSheet, {i-1, 1}, folders[i][3], over200M)
      elsif folders[i][3] >= M100 then
        worksheet_write_number(resultSheet, {i-1, 1}, folders[i][3], over100M)
      else
        worksheet_write_number(resultSheet, {i-1, 1}, folders[i][3], NULL)
      end if
      worksheet_write_string( resultSheet, {i-1, 2}, formatNb(folders[i][3]), NULL)
      worksheet_write_number( resultSheet, {i-1, 3}, folders[i][4], NULL)
      if folders[i][5] = 0 then
        worksheet_write_number(resultSheet, {i-1, 4}, folders[i][5], NULL)
      elsif folders[i][5] < year-7 then
        worksheet_write_number(resultSheet, {i-1, 4}, folders[i][5], over20G)
      elsif folders[i][5] < year-6 then
        worksheet_write_number(resultSheet, {i-1, 4}, folders[i][5], over10G)
      elsif folders[i][5] < year-5 then
        worksheet_write_number(resultSheet, {i-1, 4}, folders[i][5], over5G)
      elsif folders[i][5] < year-4 then
        worksheet_write_number(resultSheet, {i-1, 4}, folders[i][5], over2G)
      elsif folders[i][5] < year-3 then
        worksheet_write_number(resultSheet, {i-1, 4}, folders[i][5], over1G)
      elsif folders[i][5] < year-2 then
        worksheet_write_number(resultSheet, {i-1, 4}, folders[i][5], over500M)
      elsif folders[i][5] < year-1 then
        worksheet_write_number(resultSheet, {i-1, 4}, folders[i][5], over200M)
      elsif folders[i][5] < year then
        worksheet_write_number(resultSheet, {i-1, 4}, folders[i][5], over100M)
      else
        worksheet_write_number(resultSheet, {i-1, 4}, folders[i][5], NULL)
      end if
    end for
    if length(toZip) > 0 then
      worksheet_write_string(problemsSheet, {0, 0}, "Following directories or their files have exceeding full qualified name lengths.", formatBold)
      worksheet_write_string(problemsSheet, {1, 0}, "Their content might be unreachable.", formatBold)
      worksheet_write_string(problemsSheet, {2, 0}, "They must be moved or compressed.", formatBold)
      for i = 1 to length(tooLong) do
        worksheet_write_string(problemsSheet, {2+i, 0}, sprintf("%s (%d chars)", {tooLong[i][1],tooLong[i][2]}), NULL)
      end for
    end if
    workbook_close( workbook )
    setText({SB, 1}, sprintf("Finished in %d s.", {time() - t0}))
    setText({SB, 2}, sprintf("See '%s' to view results.", {outFile}))
  end if
end procedure

------------------------------------------------------------------------------

-- Application Initialization
function AppInit()
  integer lRC

  f_debug = open(InitialDir & "debug.log", "w")
  lRC = 0

  Win = createEx(Window, APPL & " v " & VERS & " - Jean-Marc DURO (" & YEAR & ")",
                 0, Default, Default, 1140, 90, 0, 0 )
  SB = createEx(StatusBar, "", Win, 0, {0.2,-1}, 0, 0, 0, 0)
  btnStart = createEx(Button, "Select directory", Win, 8, 4, 320, 23, 0, 0)

  iExw = w32Func(xLoadIcon,{instance(), "exw"} )

  closefolder = addIcon( extractIcon( resourceDir&"\\clsdfold.ico" ) )
  openfolder  = addIcon( extractIcon( resourceDir&"\\openfold.ico" ))

  location = {}

  VOID = sendMessage( Win, WM_SETICON, 0, iExw)

  setHandler(btnStart,  w32HClick,  routine_id("Click_btnStart"))

  return lRC
end function

------------------------------------------------------------------------------

-- Main
if AppInit() = 0 then
  WinMain( Win, Normal)
  close(f_debug)
end if
