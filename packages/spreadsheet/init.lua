--
-- Spreadsheet support package for SILE
-- 2023, Didier Willis
-- License: MIT
--
-- FAIRLY EXPERIMENTAL! IMPROVEMENTS WELCOME.
--
-- Table construction in SILE is compatible with package 'ptable'
-- from 3rd-party module "ptable.sile"
--
local unzipper = require("core.xutils.unzipper")
local lxp = require("lxp") -- Comes bundled with SILE

local base = require("packages.base")

local package = pl.class(base)
package._name = "spreadsheet"


local function cmd(command, options)
  return {
    command = command,
    options = options or {},
    id = "command",
    col = 0,
    lno = 0,
    pos = 0
  }
end

local function odsStartCommand (parser, command, options)
  local sheets = parser:getcallbacks().sheets
  local context = parser:getcallbacks().context

  if command == "table:table" then
    local name = options["table:name"]
    local sheet = cmd("ptable", {})
    sheets[name] = sheet
    table.insert(sheets, sheet)
    context.sheet = sheet
  elseif command == "table:table-column" then
    context.cols = context.cols or 0
    local rep = options["table:number-columns-repeated"] or 1
    context.cols = context.cols + rep
  elseif command == "table:table-row" then
    local current = context.sheet
    local row = cmd("row", {})
    local rep = tonumber(options["table:number-rows-repeated"]) or 1
    for _ = 1, rep do
      table.insert(current, row)
    end
    context.row = row
  elseif command == "table:table-cell" then
    local current = context.row
    local cell = cmd("cell", { valign="top", halign="center"})
    local rep = tonumber(options["table:number-columns-repeated"]) or 1
    for _ = 1, rep do
      table.insert(current, cell)
    end
    context.cell = cell
  end
end

local function odsEndCommand (parser, command)
  local context = parser:getcallbacks().context
  if command == "table:table-cell" then
    context.cell = nil
  elseif command == "table:table-row" then
    context.row = nil
  elseif command == "table:table" then
    local cols = context.cols
    local colspec = string.rep(string.format("%.5f%%lw", 99.9 / cols), cols, " ")
    context.sheet.options.cols = colspec
    context.sheet = nil
    context.cols = 0
  end
end

local function odsText (parser, msg)
  local context = parser:getcallbacks().context
  if context.cell then
    table.insert(context.cell, msg)
  end
end

local function parseODS (doc)
  local content = { StartElement = odsStartCommand,
              EndElement = odsEndCommand,
              CharacterData = odsText,
              _nonstrict = true,
              sheets = {},
              context = {}
        }
  local parser = lxp.new(content)
  local status, err
  if type(doc) == "string" then
    status, err = parser:parse(doc)
    if not status then return nil, err end
  else
    for element in pairs(doc) do
      status, err = parser:parse(element)
      if not status then return nil, err end
    end
  end
  status, err = parser:parse()
  if not status then return nil, err end
  parser:close()
  return content.sheets
end

local function colNumber(collAddr)
  local addr = collAddr:match("[A-Z]+")
  local m, r = 1, 0
  for i = #addr, 1, -1 do
    r = r + (string.byte(addr, i) - 64) * m
    m = m * 26
  end
  return r
end

local function ooxmlStartCommand (parser, command, options)
  local sheet = parser:getcallbacks().sheet
  local context = parser:getcallbacks().context

  if command == "row" then
    context.irow = (context.irow or 0) + 1
    local N = tonumber(options.r)
    for _ = context.irow, N-1 do
      context.irow = context.irow + 1
      table.insert(sheet, cmd("row", {}))
    end
    local row = cmd("row", {})
    table.insert(sheet, row)
    context.row = row
  elseif command == "c" then
    context.type = options.t
    local current = context.row
    context.icell = (context.icell or 0) + 1
    local N = tonumber(colNumber(options.r))
    for _ = context.icell, N-1 do
      context.icell = context.icell + 1
      table.insert(current, cmd("cell", {}))
    end
    local cell = cmd("cell", { celllid=options.r, valign="top", halign="center"})
    table.insert(current, cell)
    context.mcell = SU.max(context.mcell or 0, N-1)
    context.cell = cell
  elseif command == "v" then
    context.value = true
  end
end

local function ooxmlEndCommand (parser, command)
  local context = parser:getcallbacks().context
  if command == "c" then
    context.cell = nil
  elseif command == "row" then
    context.row = nil
    context.icell = 0
  elseif command == "v" then
    context.value = nil
  end
end

local function ooxmlText (parser, msg)
  local context = parser:getcallbacks().context
  local strings = parser:getcallbacks().strings
  if context.cell and context.value then
    if context.type == "s" then
      table.insert(context.cell, strings[tonumber(msg) + 1])
    else
      table.insert(context.cell, msg)
    end
  end
end

local function parseOOXMLWorksheet (doc, strings)
  local content = { StartElement = ooxmlStartCommand,
              EndElement = ooxmlEndCommand,
              CharacterData = ooxmlText,
              _nonstrict = true,
              sheet = cmd("ptable", {}),
              context = {},
              strings = strings or {}
        }
  local parser = lxp.new(content)
  local status, err
  if type(doc) == "string" then
    status, err = parser:parse(doc)
    if not status then return nil, err end
  else
    for element in pairs(doc) do
      status, err = parser:parse(element)
      if not status then return nil, err end
    end
  end
  status, err = parser:parse()
  if not status then return nil, err end
  parser:close()

  local cols = content.context.mcell + 1
  local colspec = string.rep(string.format("%.5f%%lw", 99.9 / cols), cols, " ")
  content.sheet.options.cols = colspec

  -- Fix number of cells on each row (i.e. for final empty cells)
  for _, row in ipairs(content.sheet) do
    for _ = #row, cols - 1 do
      table.insert(row, cmd("cell", { background="250" }))
    end
  end
  return content.sheet
end

function package:_init (options)
  base._init(self, options)
  self:loadPackage("ptable")
end

local function parseOOXMLWorkbook (doc)
  local sheets = {}
  local ids = {}
  local content = {
    StartElement = function (_, command, options)
      if command == "sheet" then
        local name = options.name
        local id = options.sheetId
        sheets[name] = "xl/worksheets/sheet"..id..".xml"
        ids[id] = sheets[name]
      end
    end,
  }
  local parser = lxp.new(content)
  local status, err
  if type(doc) == "string" then
    status, err = parser:parse(doc)
    if not status then return nil, err end
  end
  status, err = parser:parse()
  if not status then return nil, err end
  parser:close()
  return sheets, ids
end

local function parseOOXMLStrings (doc)
  local strings = {}
  local si = 0
  local content = {
    StartElement = function (_, command, _)
      if command == "si" then
        si = #strings + 1
      end
    end,
    EndElement = function (_, command, _)
      if command == "si" then
        si = 0
      end
    end,
    CharacterData = function (_, msg)
      if si > 0 then
        strings[si] = (strings[si] or "") .. msg
      end
    end
  }
  local parser = lxp.new(content)
  local status, err
  if type(doc) == "string" then
    status, err = parser:parse(doc)
    if not status then return nil, err end
  end
  status, err = parser:parse()
  if not status then return nil, err end
  parser:close()
  return strings
end

function package:_init (options)
  base._init(self, options)
  self:loadPackage("ptable")
end

function package.workbook (_, filename, sheet)
  local filepath = SILE.resolveFile(filename) or SU.error("Couldn't find file "..filename)
  local archive = unzipper(filepath) -- errors if not a zip
  local res
  local ods, _ = archive:get("content.xml")
  if ods then
    res = parseODS(ods)
    res = res[sheet] or SU.error("Spreadsheet "..sheet.." does not exist in "..filename)
  else
    local workbook, _ = archive:get("xl/workbook.xml")
    local shared = archive:get("xl/sharedStrings.xml")
    if workbook and shared then
      local sheets, ids = parseOOXMLWorkbook(workbook)
      local strings = parseOOXMLStrings(shared)
      local sheetfile
      if tonumber(sheet) then
        sheetfile = ids[""..sheet]
      else
        sheetfile = sheets[sheet]
      end
      local worksheet, _ = archive:get(sheetfile)
      if worksheet then
        res = parseOOXMLWorksheet(worksheet, strings)
      else
        SU.error("Worksheet "..sheet.. " does not exist in "..filename)
      end
    else
      SU.error("Filename '"..filename..'" is not a recognized spreadsheet')
    end
  end
  archive:dispose()
  return res
end

function package:registerCommands ()
  self:registerCommand("spreadsheet", function (options, _)
    local src = SU.required(options, "src", "spreadsheet")
    local sheet = tonumber(options.sheet) or options.sheet or 1
    local header = SU.boolean(options.header, false)
    local ptable = self:workbook(src, sheet)

    if header and #ptable > 0 then
      ptable.options.header = true
      ptable[1].options.background = "#eee"
    end
    SILE.process({ ptable })
  end)
end

package.documentation = [[
\begin{document}
The \autodoc:package{spreadsheet} experimental package provides basic tools for parsing and processing spreadsheet documents.

The supported formats are:
\begin{itemize}
\item{OpenDocument Spreadsheet Document (\code{.ods}).}
\item{Office Open XML Workbook (\code{.xlsx}).}
\end{itemize}

The \autodoc:command{\spreadsheet[src=<filename>, sheet=<name|number>, header=<boolean>]} command reads an external spreadsheet and generates a table.

The \autodoc:parameter{sheet} option defaults to 1, but may be set to another index or an explicit sheet name, for spreadsheet documents containing several sheets.

The \autodoc:parameter{header} option specifies whether first row should be considered as a header row or not in the generated table. It defaults to \code{false}.

Please note that the spreadsheets are expected to be reasonably simple: merged cells or rows are not supported, content styling is ignored, etc.
\end{document}
]]

return package
