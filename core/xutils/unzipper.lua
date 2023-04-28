--- A Lua implementation of .zip decompression.
-- Using lua-lzib, since the latter comes bundled with SILE.
--
-- License: MIT
-- Copyright 2023, Didier Willis for this adaptation.
-- Most of the ZIP parsing code was extracted and derived from:
--   https://github.com/luarocks/luarocks/blob/master/src/luarocks/tools/zip.lua
-- Proper attribution for the relevant code (also MIT-licensed):
--   Copyright 2007-2011, Kepler Project.
--   Copyright 2011-2022, the LuaRocks project authors.
--
-- BEGIN CODE ADAPTED FROM LUAROCKS
local zlib = require("zlib")

local function shr(n, m)
  return math.floor(n / 2 ^ m)
end

local function shl(n, m)
  return n * 2 ^ m
end
local function lowbits(n, m)
  return n % 2 ^ m
end

local function mode_to_windowbits(mode)
  if mode == "gzip" then
    return 31
  elseif mode == "zlib" then
    return 0
  elseif mode == "raw" then
    return -15
  end
end

-- zlib module can be provided by both lzlib and lua-lzib packages.
-- Create a compatibility layer.
local zlib_uncompress, zlib_crc32
if zlib._VERSION:match "^lua%-zlib" then
  zlib_uncompress = function(data, mode)
    return (zlib.inflate(mode_to_windowbits(mode))(data))
  end

  zlib_crc32 = function(data)
    return zlib.crc32()(data)
  end
elseif zlib._VERSION:match "^lzlib" then
  zlib_uncompress = function(data, mode)
    return zlib.decompress(data, mode_to_windowbits(mode))
  end

  zlib_crc32 = function(data)
    return zlib.crc32(zlib.crc32(), data)
  end
else
  error("unknown zlib library", 0)
end

local function number_to_lestring(number, nbytes)
  local out = {}
  for _ = 1, nbytes do
    local byte = number % 256
    table.insert(out, string.char(byte))
    number = (number - byte) / 256
  end
  return table.concat(out)
end

local function lestring_to_number(str)
  local n = 0
  local bytes = {string.byte(str, 1, #str)}
  for b = 1, #str do
    n = n + shl(bytes[b], (b - 1) * 8)
  end
  return math.floor(n)
end

local LOCAL_FILE_HEADER_SIGNATURE = number_to_lestring(0x04034b50, 4)
local CENTRAL_DIRECTORY_SIGNATURE = number_to_lestring(0x02014b50, 4)
local END_OF_CENTRAL_DIR_SIGNATURE = number_to_lestring(0x06054b50, 4)

local function ziptime_to_luatime(ztime, zdate)
  local date = {
    year = shr(zdate, 9) + 1980,
    month = shr(lowbits(zdate, 9), 5),
    day = lowbits(zdate, 5),
    hour = shr(ztime, 11),
    min = shr(lowbits(ztime, 11), 5),
    sec = lowbits(ztime, 5) * 2
  }

  if date.month == 0 then
    date.month = 1
  end
  if date.day == 0 then
    date.day = 1
  end

  return date
end

local function read_file_in_zip(zh, cdr)
  local sig = zh:read(4)
  if sig ~= LOCAL_FILE_HEADER_SIGNATURE then
    return nil, "failed reading Local File Header signature"
  end

  -- Skip over the rest of the zip file header. See
  -- zipwriter_close_file_in_zip for the format.
  zh:seek("cur", 22)
  local file_name_length = lestring_to_number(zh:read(2))
  local extra_field_length = lestring_to_number(zh:read(2))
  zh:read(file_name_length)
  zh:read(extra_field_length)

  local data = zh:read(cdr.compressed_size)

  local uncompressed
  if cdr.compression_method == 8 then
    uncompressed = zlib_uncompress(data, "raw")
  elseif cdr.compression_method == 0 then
    uncompressed = data
  else
    return nil, "unknown compression method " .. cdr.compression_method
  end

  if #uncompressed ~= cdr.uncompressed_size then
    return nil, "uncompressed size doesn't match"
  end
  if cdr.crc32 ~= zlib_crc32(uncompressed) then
    return nil, "crc32 failed (expected " .. cdr.crc32 .. ") - data: " .. uncompressed
  end

  return uncompressed
end

local function process_end_of_central_dir(zh)
  local at, err = zh:seek("end", -22)
  if not at then
    return nil, err
  end

  while true do
    local sig = zh:read(4)
    if sig == END_OF_CENTRAL_DIR_SIGNATURE then
      break
    end
    at = at - 1
    local at1, _ = zh:seek("set", at)
    if at1 ~= at then
      return nil, "Could not find End of Central Directory signature"
    end
  end

  -- number of this disk (2 bytes)
  -- number of the disk with the start of the central directory (2 bytes)
  -- total number of entries in the central directory on this disk (2 bytes)
  -- total number of entries in the central directory (2 bytes)
  zh:seek("cur", 6)
  local central_directory_entries = lestring_to_number(zh:read(2))

  -- central directory size (4 bytes)
  zh:seek("cur", 4)

  local central_directory_offset = lestring_to_number(zh:read(4))

  return central_directory_entries, central_directory_offset
end

local function process_central_dir(zh, cd_entries)
  local files = {}

  for i = 1, cd_entries do
    local sig = zh:read(4)
    if sig ~= CENTRAL_DIRECTORY_SIGNATURE then
      return nil, "failed reading Central Directory signature"
    end

    local cdr = {}
    files[i] = cdr
    cdr.version_made_by = lestring_to_number(zh:read(2))
    cdr.version_needed = lestring_to_number(zh:read(2))
    cdr.bitflag = lestring_to_number(zh:read(2))
    cdr.compression_method = lestring_to_number(zh:read(2))
    cdr.last_mod_file_time = lestring_to_number(zh:read(2))
    cdr.last_mod_file_date = lestring_to_number(zh:read(2))
    cdr.last_mod_luatime = ziptime_to_luatime(cdr.last_mod_file_time, cdr.last_mod_file_date)
    cdr.crc32 = lestring_to_number(zh:read(4))
    cdr.compressed_size = lestring_to_number(zh:read(4))
    cdr.uncompressed_size = lestring_to_number(zh:read(4))
    cdr.file_name_length = lestring_to_number(zh:read(2))
    cdr.extra_field_length = lestring_to_number(zh:read(2))
    cdr.file_comment_length = lestring_to_number(zh:read(2))
    cdr.disk_number_start = lestring_to_number(zh:read(2))
    cdr.internal_attr = lestring_to_number(zh:read(2))
    cdr.external_attr = lestring_to_number(zh:read(4))
    cdr.offset = lestring_to_number(zh:read(4))
    cdr.file_name = zh:read(cdr.file_name_length)
    cdr.extra_field = zh:read(cdr.extra_field_length)
    cdr.file_comment = zh:read(cdr.file_comment_length)
  end
  return files
end
-- END CODE ADAPTED FROM LUAROCKS

local unzipper = pl.class({})

function unzipper:_init(filename)
  local err1
  self.fd, err1 = io.open(filename, "rb")
  if not self.fd then
    SU.error("Failure to open file '"..filename.."' ("..err1..")")
  end

  local cd_entries, cd_offset = process_end_of_central_dir(self.fd)
  if not cd_entries then
    self.fd:close()
    SU.error("Failure to process ZIP file '"..filename.."' ("..cd_offset..")")
  end

  local offset, err2 = self.fd:seek("set", cd_offset)
  if not offset then
    self.fd:close()
    SU.error("Failure to process ZIP file '"..filename.."' ("..err2..")")
  end

  local files, err3 = process_central_dir(self.fd, cd_entries)
  if not files then
    self.fd:close()
    SU.error("Failure to process ZIP file '"..filename.."' ("..err3..")")
  end
  self.files = files
end

function unzipper:dispose()
  if self.fd then
    self.fd:close()
  end
end

-- Uncompress a file identified by its name from the ZIP archive.
-- Returns the inflated contents upon succees; otherwise nil and an error message.
function unzipper:get(filename)
  if not self.fd then
    return nil, "Unzipper not properly initialized"
  end
  if not self.files then
    return nil, "Unzipper not properly initialized"
  end
  for i = #self.files, 1, -1 do
    -- N.B. Loop from end, as files can be added to a ZIP over a previous entry
    local cdr = self.files[i]
    if cdr.file_name == filename then
      local ok, errseek = self.fd:seek("set", cdr.offset)
      if not ok then
        return nil, errseek
      end
      local contents, errget = read_file_in_zip(self.fd, cdr)
      if not contents then
        return nil, errget
      end
      return contents
    end
  end
  return nil, "File not found"
end

return unzipper
