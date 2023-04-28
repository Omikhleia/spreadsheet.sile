package = "spreadsheet.sile"
version = "dev-1"
source = {
  url = "git+https://github.com/Omikhleia/spreadsheet.sile.git",
}
description = {
  summary = "Spreadsheet experimental support (ODS, XSLX) for the SILE typesetting system.",
  detailed = [[
    This collection packages for the SILE typesetter system allows reading spreadsheets
    and including them as tables in your documents.
  ]],
  homepage = "https://github.com/Omikhleia/spreadsheet.sile",
  license = "MIT",
}
dependencies = {
   "lua >= 5.1",
   "ptable.sile",
}
build = {
  type = "builtin",
  modules = {
    ["sile.packages.spreadsheet"] = "packages/spreadsheet/init.lua",
    ["sile.core.xutils.unzipper"] = "core/xutils/unzipper.lua",
  }
}
