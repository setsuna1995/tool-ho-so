$ErrorActionPreference = "Stop"

$SRC = "d:\tool-ho-so"
$DEST = "d:\tool-ho-so\05. Hồ sơ - Tư vấn hiệu quả công thức sản phẩm Bánh ăn dặm VIAM (2027)"
$D1 = Join-Path $DEST "01. Hồ sơ đạo đức đề cương"
$D2 = Join-Path $DEST "02. Hồ sơ khoa học đề cương"
$D3 = Join-Path $DEST "03. Công văn mời chuyên gia"
$D4 = Join-Path $DEST "04. Hồ sơ nghiệm thu"
New-Item -ItemType Directory -Force -Path $D1,$D2,$D3,$D4 | Out-Null

$TITLE_OLD = 'Đánh giá hiệu quả sản phẩm sữa dinh dưỡng pha sẵn KUN DOCTOR COLOSTRUM lên tình trạng dinh dưỡng, miễn dịch, tiêu hóa và giấc ngủ của trẻ từ 24 đến 72 tháng tuổi'
$TITLE_NEW = 'Tư vấn hiệu quả công thức sản phẩm Bánh ăn dặm VIAM'
# PowerShell's tokenizer treats curly quote glyphs as string delimiters even inside an
# already-quoted string, so literal “ ” characters in source break parsing. Build them
# from char codes instead and always interpolate with ${...} to avoid identifier bleed
# into the following Vietnamese word.
$LDQ = [char]0x201C
$RDQ = [char]0x201D

# ---- copy source files (selectively, excluding deep-protocol / name-change / pptx / TNLS / duplicates) ----
$copies = @(
  @{ From = "01. Hồ sơ đạo đức đề cương - mẫu COLOSTRUM\00. QĐ Giao đề tài.docx"; To = "$D1\00. QĐ Giao đề tài.docx" }
  @{ From = "01. Hồ sơ đạo đức đề cương - mẫu COLOSTRUM\01. QĐTLHĐ đạo đức đề cương.docx"; To = "$D1\01. QĐTLHĐ đạo đức đề cương.docx" }
  @{ From = "01. Hồ sơ đạo đức đề cương - mẫu COLOSTRUM\02. BB họp HĐ đạo đức - KUN COLOSTRUM.docx"; To = "$D1\02. BB họp HĐ đạo đức.docx" }
  @{ From = "01. Hồ sơ đạo đức đề cương - mẫu COLOSTRUM\03. BB kiểm phiếu HĐ đạo đức.docx"; To = "$D1\03. BB kiểm phiếu HĐ đạo đức.docx" }
  @{ From = "01. Hồ sơ đạo đức đề cương - mẫu COLOSTRUM\04. Dr.Kun QĐ chấp nhận đạo đức.docx"; To = "$D1\04. QĐ chấp nhận đạo đức.docx" }
  @{ From = "01. Hồ sơ đạo đức đề cương - mẫu COLOSTRUM\Bảng kiểm đánh giá đạo đức.doc"; To = "$D1\Bảng kiểm đánh giá đạo đức.doc" }
  @{ From = "01. Hồ sơ đạo đức đề cương - mẫu COLOSTRUM\Truong_Hong_Son_ly-lich-khoa-hoc-2024.docx"; To = "$D1\Lý lịch khoa học - Trương Hồng Sơn.docx"; NoEdit = $true }

  @{ From = "02. Hồ sơ khoa học đề cương - mẫu COLOSTRUM\05. Dr.Kun QD TLHDKH đề cương.docx"; To = "$D2\05. QĐ TLHĐ khoa học xét đề cương.docx" }
  @{ From = "02. Hồ sơ khoa học đề cương - mẫu COLOSTRUM\06. Dr.Kun Bien ban hop thong qua de cuong de tai.docx"; To = "$D2\06. BB họp thông qua đề cương.docx" }
  @{ From = "02. Hồ sơ khoa học đề cương - mẫu COLOSTRUM\07. Dr.Kun Bien ban kiem phieu thong qua de cuong.docx"; To = "$D2\07. BB kiểm phiếu thông qua đề cương.docx" }
  @{ From = "02. Hồ sơ khoa học đề cương - mẫu COLOSTRUM\08. Dr.Kun QĐ phe-duyet-de-tai.docx"; To = "$D2\08. QĐ phê duyệt đề tài.docx" }
  @{ From = "02. Hồ sơ khoa học đề cương - mẫu COLOSTRUM\Dr.Kun Phieu cham diem HD de cuong.docx"; To = "$D2\Phiếu chấm điểm HĐ đề cương.docx" }
  @{ From = "02. Hồ sơ khoa học đề cương - mẫu COLOSTRUM\Dr.Kun Phieu nhan xet danh gia ho so.docx"; To = "$D2\Phiếu nhận xét đánh giá hồ sơ.docx" }

  @{ From = "03. CV mời chuyên gia - mẫu COLOSTRUM\CV mời chuyên gia.doc"; To = "$D3\Công văn mời chuyên gia.doc" }

  @{ From = "04. Hồ sơ nghiệm thu\04. Hồ sơ nghiệm thu\9. Quyết định THÀNH LẬP HĐ nghiệm thu.doc"; To = "$D4\9. Quyết định thành lập HĐ nghiệm thu.doc" }
  @{ From = "04. Hồ sơ nghiệm thu\04. Hồ sơ nghiệm thu\10. Biên bản HỌP HĐ nghiệm thu.doc"; To = "$D4\10. Biên bản họp HĐ nghiệm thu.doc" }
  @{ From = "04. Hồ sơ nghiệm thu\04. Hồ sơ nghiệm thu\11. Biên bản KIỂM PHIẾU nghiệm thu.doc"; To = "$D4\11. Biên bản kiểm phiếu nghiệm thu.doc" }
  @{ From = "04. Hồ sơ nghiệm thu\04. Hồ sơ nghiệm thu\12. Quyết định công nhận kết quả đề tài.doc"; To = "$D4\12. Quyết định công nhận kết quả đề tài.doc" }
  @{ From = "04. Hồ sơ nghiệm thu\04. Hồ sơ nghiệm thu\Phiếu CHẤM ĐIỂM nghiệm thu-(TVCT_ĐGHQ).docx"; To = "$D4\Phiếu chấm điểm nghiệm thu (TVCT_ĐGHQ).docx" }
  @{ From = "04. Hồ sơ nghiệm thu\04. Hồ sơ nghiệm thu\Phiếu ký nhận tiền.docx"; To = "$D4\Phiếu ký nhận tiền.docx" }
  @{ From = "04. Hồ sơ nghiệm thu\04. Hồ sơ nghiệm thu\Phiếu NHẬN XÉT nghiệm thu.doc"; To = "$D4\Phiếu nhận xét nghiệm thu.doc" }
)

foreach ($c in $copies) {
  Copy-Item -LiteralPath (Join-Path $SRC $c.From) -Destination $c.To -Force
}

# ---- Word COM helpers ----
$word = New-Object -ComObject Word.Application
$word.Visible = $false

# Both helpers below swallow their own errors (log + continue) on purpose: a single
# missed match or bad table index should not blow up the whole run and lose every other
# file's edits. Anything Write-Warning'd here is exactly what to go fix by hand afterward.
function Replace-Text($doc, [string]$find, [string]$replace, [bool]$wildcards = $false) {
  try {
    $rng = $doc.Content
    $rng.Find.ClearFormatting()
    $ok = $rng.Find.Execute($find, $false, $false, $wildcards, $false, $false, $true, 1, $false, $replace, 2)
    if (-not $ok) {
      $preview = $find.Substring(0, [Math]::Min(60, $find.Length))
      Write-Warning "[$($doc.Name)] Khong tim thay: '$preview...'"
    }
  } catch {
    Write-Warning "[$($doc.Name)] Loi Replace-Text: $($_.Exception.Message)"
  }
}

function Set-Cell($doc, [int]$tableIndex, [int]$row, [int]$col, [string]$text) {
  try {
    $doc.Tables.Item($tableIndex).Cell($row, $col).Range.Text = $text
  } catch {
    Write-Warning "[$($doc.Name)] Loi Set-Cell (table=$tableIndex row=$row col=$col): $($_.Exception.Message)"
  }
}

function Open-Doc($path) {
  return $word.Documents.Open($path, $false, $false)
}

function Save-Close($doc) {
  try {
    $doc.Save()
  } catch {
    Write-Warning "[$($doc.Name)] Loi khi Save: $($_.Exception.Message)"
  } finally {
    $doc.Close()
  }
}

# Everything below runs inside try/finally so a stray error never leaves a headless
# WINWORD.EXE process orphaned/locking files (that bit us earlier this session).
try {

# ---- 01. QĐ Giao đề tài ----
$doc = Open-Doc "$D1\00. QĐ Giao đề tài.docx"
Replace-Text $doc $TITLE_OLD $TITLE_NEW
Replace-Text $doc '2024' '2027'
Replace-Text $doc 'Cử nhân HOÀNG HÀ LINH^13' '' $true
Replace-Text $doc 'Cử nhân PHẠM HỒNG NGỌC^13' '' $true
Replace-Text $doc 'Cử nhân TRƯƠNG PHAN HỒNG HÀ' ''
Save-Close $doc

# ---- 02. QĐTLHĐ đạo đức đề cương ----
$doc = Open-Doc "$D1\01. QĐTLHĐ đạo đức đề cương.docx"
Replace-Text $doc $TITLE_OLD $TITLE_NEW
Replace-Text $doc '2024' '2027'
Set-Cell $doc 2 1 1 'GS. Ts. Nguyễn Công Khẩn'
Set-Cell $doc 2 1 2 'Hội đồng Đạo đức Y sinh Quốc gia'
Set-Cell $doc 2 1 3 'Chủ tịch Hội đồng'
Set-Cell $doc 2 2 1 'PGs. Ts. Nguyễn Xuân Ninh'
Set-Cell $doc 2 2 2 'Viện Y học ứng dụng Việt Nam'
Set-Cell $doc 2 2 3 'Thành viên'
Set-Cell $doc 2 3 1 'PGs. Ts. Lê Bạch Mai'
Set-Cell $doc 2 3 2 'Nguyên Phó Viện trưởng Viện Dinh dưỡng Quốc gia'
Set-Cell $doc 2 3 3 'Thành viên'
Set-Cell $doc 2 4 1 'PGs. Ts. Nguyễn Thị Lâm'
Set-Cell $doc 2 4 2 'Nguyên Phó Viện trưởng Viện Dinh dưỡng Quốc gia'
Set-Cell $doc 2 4 3 'Thành viên'
Set-Cell $doc 2 5 1 'PGs. Ts. Nguyễn Quang Dũng'
Set-Cell $doc 2 5 2 'Bộ môn Dinh dưỡng - Đại học Y Hà Nội'
Set-Cell $doc 2 5 3 'Thành viên'
Set-Cell $doc 3 1 1 'Hoàng Hà Linh'
Set-Cell $doc 3 1 2 'Viện Y học ứng dụng Việt Nam'
Set-Cell $doc 3 2 1 'Trương Phan Hồng Hà'
Set-Cell $doc 3 2 2 'Viện Y học ứng dụng Việt Nam'
Save-Close $doc

# ---- 03. BB họp HĐ đạo đức ----
$doc = Open-Doc "$D1\02. BB họp HĐ đạo đức.docx"
Replace-Text $doc 'Tên đề tài: Đánh giá hiệu quả sản phẩm sữa dinh dưỡng pha sẵn KUN DOCTOR COLOSTRUM lên tình trạng dinh dưỡng, miễn dịch, tiêu hóa và giấc ngủ của trẻ từ 24 đến 72 tháng tuổi.' "Tên đề tài: $TITLE_NEW."
Replace-Text $doc 'PGs. Ts. Hoàng Thị Thanh' 'GS. Ts. Nguyễn Công Khẩn'
# these carry the real COLOSTRUM meeting's decision number/date - must not leak into the new dossier
Replace-Text $doc 'Quyết định số: 04/QĐ-YHUD/2024 ngày 19 tháng 04 năm 2024' 'Quyết định số: ……/QĐ-YHUD/2027 ngày …… tháng …… năm 2027'
Replace-Text $doc 'Thời gian: ngày 25 tháng 04 năm 2024' 'Thời gian: ngày …… tháng …… năm 2027'
Save-Close $doc

# ---- 04. BB kiểm phiếu HĐ đạo đức ----
$doc = Open-Doc "$D1\03. BB kiểm phiếu HĐ đạo đức.docx"
Replace-Text $doc "Tên đề tài: $TITLE_OLD." "Tên đề tài: $TITLE_NEW."
Replace-Text $doc '2024' '2027'
Save-Close $doc

# ---- 05. QĐ chấp nhận đạo đức ----
$doc = Open-Doc "$D1\04. QĐ chấp nhận đạo đức.docx"
Replace-Text $doc "${LDQ}${TITLE_OLD}${RDQ}." "${LDQ}${TITLE_NEW}${RDQ}."
Replace-Text $doc 'Địa điểm triển khai nghiên cứu: tỉnh Thái Nguyên.' 'Địa điểm triển khai nghiên cứu: ……………………………….'
Replace-Text $doc 'Thời gian nghiên cứu: Từ 12/2024 đến 05/2024' 'Thời gian nghiên cứu: Từ 01/2027 đến 12/2027'
Replace-Text $doc '2024' '2027'
Set-Cell $doc 2 1 2 "CHỦ TỊCH HỘI ĐỒNG`rGS. Ts. Nguyễn Công Khẩn"
Save-Close $doc

# ---- Bảng kiểm đánh giá đạo đức ----
$doc = Open-Doc "$D1\Bảng kiểm đánh giá đạo đức.doc"
Replace-Text $doc "Tên nghiên cứu: $TITLE_OLD." "Tên nghiên cứu: $TITLE_NEW."
Replace-Text $doc 'Ngày       tháng       năm 2024' 'Ngày       tháng       năm 2027'
Save-Close $doc

# ---- 05. QĐ TLHĐ khoa học xét đề cương ----
$doc = Open-Doc "$D2\05. QĐ TLHĐ khoa học xét đề cương.docx"
Replace-Text $doc $TITLE_OLD $TITLE_NEW
Replace-Text $doc '2024' '2027'
Set-Cell $doc 2 1 1 'Gs. Ts. Nguyễn Công Khẩn'
Set-Cell $doc 2 1 2 'Hội đồng Đạo đức Y sinh Quốc gia'
Set-Cell $doc 2 1 3 'Chủ tịch Hội đồng'
Set-Cell $doc 2 2 1 'PGs. Ts. Nguyễn Xuân Ninh'
Set-Cell $doc 2 2 2 'Viện Y học ứng dụng Việt Nam'
Set-Cell $doc 2 2 3 'Phản biện 1'
Set-Cell $doc 2 3 1 'PGs. Ts. Trần Quang Trung'
Set-Cell $doc 2 3 2 'Hiệp hội Sữa Việt Nam'
Set-Cell $doc 2 3 3 'Phản biện 2'
Set-Cell $doc 2 4 1 'PGs. Ts. Lê Bạch Mai'
Set-Cell $doc 2 4 2 'Nguyên Phó Viện trưởng Viện Dinh dưỡng Quốc gia'
Set-Cell $doc 2 4 3 'Ủy viên'
Set-Cell $doc 2 5 1 'PGs. Ts. Nguyễn Thị Lâm'
Set-Cell $doc 2 5 2 'Nguyên Phó Viện trưởng Viện Dinh dưỡng Quốc gia'
Set-Cell $doc 2 5 3 'Ủy viên'
Set-Cell $doc 3 1 2 'Hoàng Hà Linh'
Set-Cell $doc 3 1 3 'Viện Y học ứng dụng Việt Nam'
Set-Cell $doc 3 2 2 'Trương Phan Hồng Hà'
Set-Cell $doc 3 2 3 'Viện Y học ứng dụng Việt Nam'
Save-Close $doc

# ---- 06. BB họp thông qua đề cương ----
$doc = Open-Doc "$D2\06. BB họp thông qua đề cương.docx"
Replace-Text $doc "1. Tên đề tài nghiên cứu khoa học: $TITLE_OLD." "1. Tên đề tài nghiên cứu khoa học: $TITLE_NEW."
Replace-Text $doc '2024' '2027'
Replace-Text $doc 'PGs. Ts. Hoàng Thị Thanh - Chủ tịch Hội đồng điều khiển phiên họp' 'Gs. Ts. Nguyễn Công Khẩn - Chủ tịch Hội đồng điều khiển phiên họp'
Save-Close $doc

# ---- 07. BB kiểm phiếu thông qua đề cương ----
$doc = Open-Doc "$D2\07. BB kiểm phiếu thông qua đề cương.docx"
Replace-Text $doc "Tên đề tài: $TITLE_OLD." "Tên đề tài: $TITLE_NEW."
Replace-Text $doc '2024' '2027'
Save-Close $doc

# ---- 08. QĐ phê duyệt đề tài ----
$doc = Open-Doc "$D2\08. QĐ phê duyệt đề tài.docx"
Replace-Text $doc "${LDQ}${TITLE_OLD}${RDQ}." "${LDQ}${TITLE_NEW}${RDQ}."
Replace-Text $doc 'Thời gian thực hiện của đề tài: từ tháng 12/2024 đến tháng 05/2025' 'Thời gian thực hiện của đề tài: từ tháng 01/2027 đến tháng 12/2027'
Replace-Text $doc '2025' '2027'
Replace-Text $doc '2024' '2027'
Save-Close $doc

# ---- Phiếu chấm điểm HĐ đề cương ----
$doc = Open-Doc "$D2\Phiếu chấm điểm HĐ đề cương.docx"
Replace-Text $doc "1. Tên Đề tài: $TITLE_OLD." "1. Tên Đề tài: $TITLE_NEW."
Replace-Text $doc '2024' '2027'
Save-Close $doc

# ---- Phiếu nhận xét đánh giá hồ sơ ----
$doc = Open-Doc "$D2\Phiếu nhận xét đánh giá hồ sơ.docx"
Replace-Text $doc "Tên đề tài: $TITLE_OLD." "Tên đề tài: $TITLE_NEW."
Replace-Text $doc '2024' '2027'
Save-Close $doc

# ---- Công văn mời chuyên gia ----
$doc = Open-Doc "$D3\Công văn mời chuyên gia.doc"
# The original intro is one 446-char paragraph of COLOSTRUM-specific marketing/rationale
# text - Word's Find.Text property hard-caps search strings at ~255 chars, and we won't
# fabricate replacement rationale for a product we have no data on. Strip it in three
# shorter passes (each under the limit) and leave a bracketed TODO for a human to fill in.
# All specific/long-text replacements run BEFORE the blanket '2024'->'2027' below, since
# that blanket pass would otherwise consume the '2024' inside these longer search strings
# and make them silently stop matching (bit us on 3 other files this run).
Replace-Text $doc 'Suy dinh dưỡng ở trẻ em dưới 5 tuổi – đặc biệt là suy dinh dưỡng thấp còi vẫn là một vấn đề sức khỏe cộng đồng.' '[Bổ sung bối cảnh/lý do triển khai dự án tại đây]'
Replace-Text $doc 'Một trong những giải pháp làm giảm tình trạng suy dinh dưỡng ở trẻ dưới 5 tuổi là sử dụng các sản phẩm bổ sung dinh dưỡng trong hệ thống trường mầm non.' ''
Replace-Text $doc 'Nhằm đánh giá tình trạng suy dinh dưỡng ở trẻ dưới 5 tuổi và hiệu quả của sản phẩm bổ sung dinh dưỡng LOF KUN COLOSTRUM, Viện Y học ứng dụng Việt Nam tiến hành triển khai nghiên cứu' 'Viện Y học ứng dụng Việt Nam tiến hành triển khai đề tài'
Replace-Text $doc "${LDQ}${TITLE_OLD}${RDQ}." "${LDQ}${TITLE_NEW}${RDQ}."
Replace-Text $doc 'Nghiên cứu được triển khai trong 06 tháng, trong đó thời gian can thiệp là 04 tháng.' 'Thời gian thực hiện dự kiến: 01/2027 đến 12/2027.'
Replace-Text $doc 'Thời gian: 9 giờ 00 – sáng thứ 7 ngày 07 tháng 12 năm 2024.' 'Thời gian: …… giờ ……, ngày …… tháng …… năm 2027.'
Replace-Text $doc '2024' '2027'
Save-Close $doc

# ---- 9. QĐ thành lập HĐ nghiệm thu ----
$doc = Open-Doc "$D4\9. Quyết định thành lập HĐ nghiệm thu.doc"
Replace-Text $doc '20xx' '2027'
Replace-Text $doc "${LDQ}Tên đề tài${RDQ}" "${LDQ}${TITLE_NEW}${RDQ}"
Set-Cell $doc 2 1 2 'PGs.Ts. Phạm Văn Hoàn'
Set-Cell $doc 2 1 3 'Viện Y học ứng dụng Việt Nam'
Set-Cell $doc 2 1 4 "Chủ tịch`rHội đồng"
Set-Cell $doc 2 2 2 'Gs. Ts. Nguyễn Công Khẩn'
Set-Cell $doc 2 2 3 'Hội đồng Đạo đức Y sinh Quốc gia'
Set-Cell $doc 2 2 4 'Phản biện 1'
Set-Cell $doc 2 3 2 'PGs. Ts. Nguyễn Thị Lâm'
Set-Cell $doc 2 3 3 'Nguyên Phó Viện trưởng Viện Dinh dưỡng Quốc gia'
Set-Cell $doc 2 3 4 'Phản biện 2'
Set-Cell $doc 2 4 2 'PGs. Ts. Ninh Thị Nhung'
Set-Cell $doc 2 4 3 'Trường Đại học Y Dược Thái Bình'
Set-Cell $doc 2 4 4 'Ủy viên'
Set-Cell $doc 2 5 2 'PGs.Ts. Trần Văn Ơn'
Set-Cell $doc 2 5 3 'Trường Đại học Dược Hà Nội'
Set-Cell $doc 2 5 4 'Uỷ viên'
Set-Cell $doc 3 1 1 '1. Hoàng Hà Linh'
Set-Cell $doc 3 1 2 'Viện Y học ứng dụng Việt Nam'
Set-Cell $doc 3 2 1 '2. Trương Phan Hồng Hà'
Set-Cell $doc 3 2 2 'Viện Y học ứng dụng Việt Nam'
Save-Close $doc

# ---- 10. Biên bản họp HĐ nghiệm thu ----
$doc = Open-Doc "$D4\10. Biên bản họp HĐ nghiệm thu.doc"
Replace-Text $doc '20xx' '2027'
Replace-Text $doc "1. Tên đề tài: Tên đề tài" "1. Tên đề tài: $TITLE_NEW"
Replace-Text $doc 'Chủ nhiệm đề tài: Tên 1' 'Chủ nhiệm đề tài: Trương Hồng Sơn'
Replace-Text $doc 'Đồng chủ nhiệm đề tài: Tên 2' 'Đồng chủ nhiệm đề tài: '
Replace-Text $doc 'Tên 3 - Chủ tịch Hội đồng điều khiển phiên họp' 'Phạm Văn Hoàn - Chủ tịch Hội đồng điều khiển phiên họp'
Replace-Text $doc '5. Số thành viên Hội đồng theo quyết định là …… người' '5. Số thành viên Hội đồng theo quyết định là 05 người'
Save-Close $doc

# ---- 11. Biên bản kiểm phiếu nghiệm thu ----
$doc = Open-Doc "$D4\11. Biên bản kiểm phiếu nghiệm thu.doc"
Replace-Text $doc '20xx' '2027'
Replace-Text $doc "1. Tên đề tài: Tên đề tài" "1. Tên đề tài: $TITLE_NEW"
Replace-Text $doc 'Chủ nhiệm đề tài: Tên 1' 'Chủ nhiệm đề tài: Trương Hồng Sơn'
Replace-Text $doc 'Đồng chủ nhiệm đề tài: Tên 2' 'Đồng chủ nhiệm đề tài: '
Save-Close $doc

# ---- 12. QĐ công nhận kết quả đề tài ----
$doc = Open-Doc "$D4\12. Quyết định công nhận kết quả đề tài.doc"
Replace-Text $doc '20xx' '2027'
Replace-Text $doc '"Tên đề tài"' "${LDQ}${TITLE_NEW}${RDQ}"
Save-Close $doc

# ---- Phiếu chấm điểm nghiệm thu (TVCT_ĐGHQ) ----
$doc = Open-Doc "$D4\Phiếu chấm điểm nghiệm thu (TVCT_ĐGHQ).docx"
Replace-Text $doc "1. Tên đề tài: Tên đề tài" "1. Tên đề tài: $TITLE_NEW"
Replace-Text $doc 'Chủ nhiệm đề tài: Tên 1' 'Chủ nhiệm đề tài: Trương Hồng Sơn'
Save-Close $doc

# ---- Phiếu ký nhận tiền ----
$doc = Open-Doc "$D4\Phiếu ký nhận tiền.docx"
Replace-Text $doc "${LDQ}Đánh giá hiệu quả sản phẩm thực phẩm chức năng Viên nang Đông trùng hạ thảo CordySen${RDQ}" "${LDQ}${TITLE_NEW}${RDQ}"
Set-Cell $doc 2 7 2 'Hoàng Hà Linh'
Save-Close $doc

# ---- Phiếu nhận xét nghiệm thu ----
$doc = Open-Doc "$D4\Phiếu nhận xét nghiệm thu.doc"
Replace-Text $doc '20xx' '2027'
Replace-Text $doc "Tên đề tài: Tên đề tài" "Tên đề tài: $TITLE_NEW"
Replace-Text $doc 'Chủ nhiệm: Tên 1' 'Chủ nhiệm: Trương Hồng Sơn'
Replace-Text $doc 'Đồng chủ nhiệm: Tên 2' 'Đồng chủ nhiệm: '
Save-Close $doc

} finally {
  $word.Quit()
}
Write-Output "DONE"
