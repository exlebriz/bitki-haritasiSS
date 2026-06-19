# =====================================================================
#  BİTKİ NUMUNELEME İSTASYONLARI  ·  PhysIQ Dynamic
#  İnteraktif kayıt + görselleştirme platformu
#  - Tür filtresi, açılır/kapanır lejant
#  - GPS / konum bulma + enlem-boylam gösterimi
#  - Konuma yeni kayıt ekleme (tür otomatik tamamlama + manuel)
#  - Yanlış kayıt silme
#  - Tarihsel sürüm geçmişi (son 20 sürüm) + geri dönme
#  - Excel ile toplu yükleme / dışa aktarma
#
#  Yükleme: https://connect.posit.cloud/cm-yesilkanat
# =====================================================================

library(shiny)
library(bslib)
library(bsicons)
library(leaflet)
library(leaflet.extras)
library(readxl)
library(writexl)
library(DT)
library(dplyr)
library(shinyWidgets)
library(htmlwidgets)

`%||%` <- function(a, b) if (is.null(a) || length(a) == 0 || (length(a) == 1 && is.na(a))) b else a

# ---------------------------------------------------------------------
# 0. Sabitler / yardımcılar
# ---------------------------------------------------------------------
DEFAULT_DATA_PATH <- file.path("data", "istasyonlar.xlsx")
MAX_VERSIONS <- 20

PALETTE <- c(
  "#2e7d32", "#c62828", "#1565c0", "#ef6c00", "#6a1b9a",
  "#00838f", "#ad1457", "#558b2f", "#4527a0", "#d84315",
  "#00695c", "#9e9d24", "#283593", "#bf360c", "#37474f",
  "#7b1fa2", "#0277bd", "#f9a825", "#5d4037", "#c2185b"
)

# Veri ekleme listesi için sabit tür referansı (Bilimsel · Türkçe · 3 harfli kod)
SPECIES_REF <- data.frame(
  Bilimsel = c(
    "Cerasus avium", "Crataegus monogyna", "Crataegus pentagyna", "Crataegus pontica",
    "Crataegus orientalis", "Cistus creticus", "Cistus salviifolius", "Cornus mas",
    "Cydonia oblonga", "Diospyros kaki", "Diospyros lotus", "Hypericum perforatum",
    "Hypericum orientale", "Juniperus oxycedrus", "Juglans regia", "Mespilus germanica",
    "Morus alba", "Morus nigra", "Origanum rotundifolium", "Origanum vulgare",
    "Punica granatum", "Prunus divaricata", "Rosa canina", "Rubus caesius",
    "Rubus caucasicus", "Rubus saxatilis", "Rubus idaeus", "Rhus coriaria",
    "Satureja spicigera", "Satureja hortensis", "Thymus fallax", "Thymus pubescens",
    "Thymus praecox", "Thymus transcaucasicus", "Tilia dasystyla subsp. caucasica",
    "Vaccinium arctostaphylos", "Vaccinium myrtillus", "Viburnum opulus",
    "Laurocerasus officinalis"),
  Turkce = c(
    "Ku\u015f kiraz\u0131", "Adi al\u0131\u00e7", "Be\u015f \u00e7ekirdekli al\u0131\u00e7", "Pontik al\u0131\u00e7",
    "Do\u011fu al\u0131c\u0131", "Pembe laden", "Beyaz laden", "K\u0131z\u0131lc\u0131k",
    "Ayva", "Trabzon hurmas\u0131", "Yabani hurma", "Sar\u0131 kantaron",
    "Do\u011fu kantaronu", "Katran ard\u0131c\u0131", "Ceviz", "Mu\u015fmula",
    "Ak dut", "Kara dut", "Yuvarlak yaprakl\u0131 kekik", "Adi kekik",
    "Nar", "Yabani erik", "Ku\u015fburnu", "Mavi b\u00f6\u011f\u00fcrtlen",
    "Kafkas b\u00f6\u011f\u00fcrtleni", "Ta\u015f b\u00f6\u011f\u00fcrtleni", "Ahududu", "Sumak",
    "Sivri kekik", "Yazl\u0131k sater", "Kekik (T. fallax)", "Kekik (T. pubescens)",
    "Mor kekik", "Kafkas keki\u011fi", "Kafkas \u0131hlamuru",
    "Likapa", "Yaban mersini", "Gilaburu",
    "Karayemi\u015f"),
  Kod = c(
    "CAV", "CMO", "CPE", "CPO", "COR", "CCR", "CSA", "CMA", "CYO", "DKA",
    "DLO", "HPE", "HOR", "JOX", "JRE", "MGE", "MAL", "MNI", "ORO", "OVU",
    "PGR", "PDI", "RCA", "RCE", "RCC", "RSA", "RID", "RHU", "SSP", "SHO",
    "TFA", "TPU", "THP", "THT", "TDA", "VAR", "VMY", "VOP", "LOF"),
  stringsAsFactors = FALSE)
SPECIES_REF$Turkce <- enc2utf8(SPECIES_REF$Turkce)

# Yazılabilir geçmiş dosyası yolu (önce app klasörü, olmazsa geçici dizin)
history_path <- local({
  cands <- c(file.path("data", "history.rds"),
             file.path(tempdir(), "bitki_history.rds"))
  for (p in cands) {
    ok <- tryCatch({
      d <- dirname(p); if (!dir.exists(d)) dir.create(d, recursive = TRUE)
      tf <- file.path(d, ".wtest"); file.create(tf); unlink(tf); TRUE
    }, error = function(e) FALSE)
    if (ok) return(p)
  }
  file.path(tempdir(), "bitki_history.rds")
})

# Yazılabilir geçmiş KLASÖRÜ (her sürüm ayrı .xlsx; son 20 dosya saklanır)
history_dir <- local({
  cands <- c(file.path("data", "history"), file.path(tempdir(), "bitki_history"))
  for (p in cands) {
    ok <- tryCatch({
      if (!dir.exists(p)) dir.create(p, recursive = TRUE)
      tf <- file.path(p, ".wtest"); file.create(tf); unlink(tf); TRUE
    }, error = function(e) FALSE)
    if (ok) return(p)
  }
  d <- file.path(tempdir(), "bitki_history"); dir.create(d, showWarnings = FALSE); d
})

# Sayısal temizleyici (—, °, virgül vb.)
num <- function(x) {
  x <- as.character(x); x <- gsub(",", ".", x, fixed = TRUE)
  x <- gsub("[^0-9.-]", "", x); suppressWarnings(as.numeric(x))
}

# Tarih -> görüntü metni (yıl / tam tarih / Excel seri no destekli)
parse_tarih <- function(x) {
  x <- trimws(as.character(x)); out <- rep(NA_character_, length(x))
  iso <- grepl("^[0-9]{4}-[0-9]{2}-[0-9]{2}", x); out[iso] <- substr(x[iso], 1, 10)
  yr  <- grepl("^[0-9]{4}$", x) & is.na(out); out[yr] <- x[yr]
  ser <- grepl("^[0-9]{5}(\\.[0-9]+)?$", x) & is.na(out)
  out[ser] <- as.character(as.Date(round(as.numeric(x[ser])), origin = "1899-12-30"))
  rest <- is.na(out) & nzchar(x) & x != "NA"; out[rest] <- x[rest]
  out
}

# Sıralama için tarih anahtarı (yıl-yalnız ve ISO tarih güvenli)
date_key <- function(x) {
  x <- as.character(x); d <- as.Date(rep(NA_character_, length(x)))
  yr  <- grepl("^[0-9]{4}$", x);                 d[yr]  <- as.Date(paste0(x[yr], "-12-31"))
  iso <- grepl("^[0-9]{4}-[0-9]{2}-[0-9]{2}", x); d[iso] <- as.Date(substr(x[iso], 1, 10))
  d
}

# Bir veri çerçevesini standart şemaya çevir (ID, ID2, Tarih, Tür, Bilimsel, Türkçe, Enlem, Boylam)
parse_plant_data <- function(df) {
  if (is.null(df) || nrow(df) == 0) stop("Veri boş görünüyor.")
  nm <- names(df)
  low <- tolower(chartr("I\u0130", "ii", trimws(nm)))   # Türkçe I/İ tuzağına karşı
  pick <- function(pats, fb) {
    for (p in pats) { h <- which(grepl(p, low)); if (length(h)) return(h[1]) }
    if (!is.na(fb) && fb <= length(nm)) fb else NA_integer_
  }
  i_id  <- pick(c("^id$"), 1)
  i_id2 <- pick(c("^id2$", "id2", "istasyon"), NA)
  i_tar <- pick(c("tarih", "date", "y\u0131l", "yil", "year"), NA)
  i_tur <- pick(c("^t\u00fcr$", "^tur$", "code", "kisaltma"), NA)
  i_bil <- pick(c("bilim", "scientific", "latin", "species"), NA)
  i_trk <- pick(c("t\u00fcrk\u00e7e", "turkce", "turkish", "yerel"), NA)
  i_lat <- pick(c("enlem", "lat", "n\\)"), NA)
  i_lon <- pick(c("boylam", "lon", "lng", "e\\)"), NA)
  if (any(is.na(c(i_lat, i_lon))))
    stop("Enlem/Boylam s\u00fctunlar\u0131 bulunamad\u0131. L\u00fctfen \u015fablonu kullan\u0131n.")
  idv <- if (!is.na(i_id)) as.character(df[[i_id]]) else as.character(seq_len(nrow(df)))
  out <- data.frame(
    ID       = idv,
    ID2      = if (!is.na(i_id2)) as.character(df[[i_id2]]) else idv,
    Tarih    = if (!is.na(i_tar)) parse_tarih(df[[i_tar]])  else NA_character_,
    Tur      = if (!is.na(i_tur)) as.character(df[[i_tur]]) else "\u2014",
    Bilimsel = if (!is.na(i_bil)) as.character(df[[i_bil]]) else "Belirtilmemi\u015f",
    Turkce   = if (!is.na(i_trk)) as.character(df[[i_trk]]) else "",
    Enlem    = num(df[[i_lat]]),
    Boylam   = num(df[[i_lon]]),
    stringsAsFactors = FALSE)
  out$Bilimsel[is.na(out$Bilimsel) | out$Bilimsel == ""] <- "Belirtilmemi\u015f"
  out$Tur[is.na(out$Tur) | out$Tur == ""] <- "\u2014"
  out$Turkce[is.na(out$Turkce)] <- ""
  out$ID2[is.na(out$ID2) | out$ID2 == "" | out$ID2 == "NA"] <- out$ID[is.na(out$ID2) | out$ID2 == "" | out$ID2 == "NA"]
  for (cc in c("ID", "ID2", "Tarih", "Tur", "Bilimsel", "Turkce"))
    out[[cc]] <- enc2utf8(as.character(out[[cc]]))
  out
}

build_palette <- function(species) {
  sp <- sort(unique(species)); setNames(rep(PALETTE, length.out = length(sp)), sp)
}

# Bir kimlik vektöründeki en büyük sayıyı bulup +1 döndürür (yoksa 1)
next_seq <- function(vals) {
  n <- suppressWarnings(as.integer(gsub("[^0-9]", "", as.character(vals))))
  m <- suppressWarnings(max(n[is.finite(n)], na.rm = TRUE)); if (!is.finite(m)) m <- 0
  m + 1L
}

load_default <- function() {
  if (file.exists(DEFAULT_DATA_PATH)) {
    raw <- tryCatch(readxl::read_excel(DEFAULT_DATA_PATH, col_types = "text"),
                    error = function(e) NULL)
    if (!is.null(raw)) return(parse_plant_data(as.data.frame(raw)))
  }
  data.frame(ID = character(), ID2 = character(), Tarih = character(), Tur = character(),
             Bilimsel = character(), Turkce = character(),
             Enlem = numeric(), Boylam = numeric(), stringsAsFactors = FALSE)
}

# ---------------------------------------------------------------------
# 1. Tema
# ---------------------------------------------------------------------
app_theme <- bs_theme(
  version = 5, bg = "#f6f8f4", fg = "#1b2a1f",
  primary = "#2e7d32", secondary = "#00695c", success = "#558b2f",
  base_font = font_google("Inter"), heading_font = font_google("Inter Tight"),
  "border-radius" = "0.65rem")

# ---------------------------------------------------------------------
# 2. UI
# ---------------------------------------------------------------------
ui <- page_navbar(
  title = tagList(
    tags$span(bsicons::bs_icon("flower2"), style = "color:#2e7d32;"),
    tags$span("Bitki Numuneleme \u0130stasyonlar\u0131",
              style = "font-weight:700; margin-left:.35rem;")),
  theme = app_theme,
  window_title = "Bitki Numuneleme \u0130stasyonlar\u0131",
  id = "ana",

  header = tags$head(
    tags$meta(name = "viewport",
              content = "width=device-width, initial-scale=1, maximum-scale=1"),
    tags$style(HTML("
      .app-side h6 { letter-spacing:.04em; text-transform:uppercase;
        font-size:.72rem; color:#5a6b5c; margin-bottom:.4rem; margin-top:1rem; }
      .value-box .value-box-title { font-weight:600; }
      .navbar { box-shadow:0 2px 10px rgba(0,0,0,.06); }
      .help-note { font-size:.78rem; color:#6b7a6d; line-height:1.35; }
      .badge-soft { background:#e3efe1; color:#2e7d32; border-radius:1rem;
        padding:.15rem .6rem; font-size:.72rem; font-weight:600; }
      .coord-box { background:#eef4ec; border:1px solid #d7e3d4; border-radius:.5rem;
        padding:.5rem .65rem; font-family:Consolas,monospace; font-size:.86rem;
        color:#1f3324; }
      #harita { height: 70vh !important; min-height: 340px; width:100% !important; }
      #harita .leaflet-container { height:100% !important; width:100% !important;
        border-radius:.65rem; }
      @media (min-width: 992px){ #harita { height: calc(100vh - 250px) !important;
        min-height: 480px; } }
      @media (max-width: 991.98px){ .navbar-brand { font-size:.92rem; }
        .app-side { width:88vw !important; } }
      /* Açılır/kapanır lejant */
      .lejant { background:#fff; border-radius:.5rem; box-shadow:0 1px 6px rgba(0,0,0,.18);
        font-size:.78rem; max-width:230px; }
      .lejant > summary { cursor:pointer; list-style:none; padding:.4rem .6rem;
        font-weight:600; color:#2e7d32; user-select:none; }
      .lejant > summary::-webkit-details-marker { display:none; }
      .lejant > summary::before { content:'\u25B8 '; }
      .lejant[open] > summary::before { content:'\u25BE '; }
      .lejant .lej-body { padding:.1rem .6rem .5rem; max-height:42vh; overflow:auto; }
      .lejant .lej-row { display:flex; align-items:center; gap:.4rem; margin:.12rem 0; }
      .lejant .dot { width:11px; height:11px; border-radius:50%; flex:0 0 auto;
        border:1px solid #fff; box-shadow:0 0 0 1px rgba(0,0,0,.15); }
    ")),
    tags$script(HTML("
      function nudgeMap(){ setTimeout(function(){
        window.dispatchEvent(new Event('resize')); }, 350); }
      document.addEventListener('shiny:connected', nudgeMap);
      window.addEventListener('orientationchange', nudgeMap);
      document.addEventListener('shiny:value', function(e){ if(e.name==='harita') nudgeMap(); });
      document.addEventListener('click', function(e){
        if(e.target.closest('.collapse-toggle, [data-bs-toggle], .bslib-sidebar-layout button'))
          nudgeMap(); });
      Shiny.addCustomMessageHandler('getloc', function(x){
        if(navigator.geolocation){
          navigator.geolocation.getCurrentPosition(
            function(p){ Shiny.setInputValue('konum_gps',
              {lat:p.coords.latitude, lng:p.coords.longitude, t:Date.now()}); },
            function(err){ Shiny.setInputValue('konum_gps_err', (err.message||'hata')+' #'+Date.now()); },
            {enableHighAccuracy:true, timeout:10000, maximumAge:0});
        } else { Shiny.setInputValue('konum_gps_err','tarayici-desteklemiyor #'+Date.now()); }
      });
    "))
  ),

  # ============ HARİTA ============
  nav_panel(
    title = tagList(bsicons::bs_icon("geo-alt-fill"), " Harita"),
    layout_sidebar(
      sidebar = sidebar(
        class = "app-side", width = 340, open = "open", title = "Kontrol Paneli",

        tags$h6("G\u00f6r\u00fcnt\u00fclenecek t\u00fcrler"),
        pickerInput("secili_turler", NULL, choices = NULL, multiple = TRUE,
          options = pickerOptions(actionsBox = TRUE, liveSearch = TRUE,
            selectAllText = "T\u00fcm\u00fcn\u00fc se\u00e7", deselectAllText = "Temizle",
            noneSelectedText = "T\u00fcr se\u00e7iniz...",
            countSelectedText = "{0} t\u00fcr se\u00e7ili", selectedTextFormat = "count > 2")),

        tags$h6("G\u00f6r\u00fcn\u00fcm"),
        materialSwitch("lejant_goster", "T\u00fcr lejant\u0131n\u0131 g\u00f6ster",
                       value = TRUE, status = "success", right = TRUE),
        materialSwitch("etiket", "\u0130stasyon etiketleri (kal\u0131c\u0131)",
                       value = FALSE, status = "success", right = TRUE),
        materialSwitch("kumelele", "\u0130\u015faretleri k\u00fcmele",
                       value = FALSE, status = "success", right = TRUE),

        tags$hr(),
        tags$h6("Konum"),
        div(class = "help-note", style = "margin-bottom:.4rem;",
            "Konumunuzu bulun veya haritaya dokunarak bir nokta se\u00e7in."),
        uiOutput("konum_kutu"),
        div(style = "height:.45rem;"),
        actionButton("konum_bul", "Konumumu bul", icon = icon("location-crosshairs"),
                     class = "btn-sm btn-outline-secondary w-100"),
        div(style = "height:.35rem;"),
        actionButton("kayit_ekle_btn", "Bu konuma kay\u0131t ekle", icon = icon("plus"),
                     class = "btn-sm btn-success w-100"),

        tags$hr(),
        tags$h6("Toplu veri y\u00fckleme (Excel)"),
        fileInput("dosya", NULL, accept = c(".xlsx", ".xls"),
                  buttonLabel = "Se\u00e7...", placeholder = ".xlsx dosyas\u0131"),
        radioButtons("yukleme_modu", NULL,
          choices = c("Mevcut verinin yerine koy" = "replace",
                      "Mevcut veriye ekle" = "append"), selected = "replace"),
        downloadButton("sablon", "\u015eablonu indir",
                       class = "btn-sm btn-outline-secondary w-100"),
        div(style = "height:.4rem;"),
        actionButton("sifirla", "Ba\u015flang\u0131\u00e7 verisine d\u00f6n",
          icon = icon("rotate-left"), class = "btn-sm btn-outline-success w-100")
      ),

      layout_columns(fill = FALSE, gap = "0.6rem",
        col_widths = breakpoints(xs = c(6, 6, 6, 6), md = c(3, 3, 3, 3)),
        value_box("\u0130stasyon (g\u00f6r\u00fcn\u00fcr)", textOutput("vb_istasyon", inline = TRUE),
                  showcase = bsicons::bs_icon("pin-map-fill"), theme = "success"),
        value_box("Bitki t\u00fcr\u00fc", textOutput("vb_tur", inline = TRUE),
                  showcase = bsicons::bs_icon("flower1"), theme = "secondary"),
        value_box("Toplam kay\u0131t", textOutput("vb_toplam", inline = TRUE),
                  showcase = bsicons::bs_icon("database-fill"), theme = "primary"),
        value_box("Son g\u00fcncelleme", textOutput("vb_guncel", inline = TRUE),
                  showcase = bsicons::bs_icon("clock-history"), theme = "warning")),
      card(full_screen = TRUE,
        card_body(padding = 0, leafletOutput("harita", height = "100%")))
    )
  ),

  # ============ VERİ TABLOSU ============
  nav_panel(
    title = tagList(bsicons::bs_icon("table"), " Veri Tablosu"),
    card(
      card_header(div(style = "display:flex;justify-content:space-between;align-items:center;gap:.5rem;flex-wrap:wrap;",
        tags$span("Aktif s\u00fcr\u00fcm \u00b7 en yeni kay\u0131tlar \u00fcstte"),
        div(actionButton("sil_btn", "Se\u00e7ili kayd\u0131 sil", icon = icon("trash"),
              class = "btn-sm btn-outline-danger"),
            downloadButton("indir_excel", "Excel indir", class = "btn-sm btn-success")))),
      card_body(DT::DTOutput("tablo"))
    )
  ),

  # ============ ÖZET ============
  nav_panel(
    title = tagList(bsicons::bs_icon("bar-chart-fill"), " \u00d6zet"),
    layout_columns(col_widths = c(5, 7),
      card(card_header("T\u00fcr da\u011f\u0131l\u0131m\u0131"), card_body(DT::DTOutput("ozet_tablo"))),
      card(card_header("T\u00fcr ba\u015f\u0131na istasyon say\u0131s\u0131"),
           card_body(plotOutput("ozet_grafik", height = "520px"))))
  ),

  # ============ SÜRÜM GEÇMİŞİ ============
  nav_panel(
    title = tagList(bsicons::bs_icon("clock-history"), " S\u00fcr\u00fcm Ge\u00e7mi\u015fi"),
    card(
      card_header(div(style = "display:flex;justify-content:space-between;align-items:center;gap:.5rem;flex-wrap:wrap;",
        tags$span("Her ekleme/silme/y\u00fckleme bir s\u00fcr\u00fcm olarak saklan\u0131r (son 20)."),
        div(actionButton("geri_don", "Se\u00e7ili s\u00fcr\u00fcme d\u00f6n", icon = icon("clock-rotate-left"),
              class = "btn-sm btn-outline-primary"),
            downloadButton("indir_surum", "Se\u00e7ili s\u00fcr\u00fcm (.xlsx)", class = "btn-sm btn-outline-success"),
            downloadButton("indir_gecmis", "T\u00fcm ge\u00e7mi\u015f (.xlsx)", class = "btn-sm btn-success")))),
      card_body(
        div(class = "help-note", style = "margin-bottom:.5rem;",
            "Aktif s\u00fcr\u00fcm yukar\u0131da i\u015faretlidir. Kal\u0131c\u0131 yedek i\u00e7in 'T\u00fcm ge\u00e7mi\u015f' dosyas\u0131n\u0131 indirin."),
        DT::DTOutput("surum_tablo"))
    )
  ),

  # ============ HAKKINDA ============
  nav_panel(
    title = tagList(bsicons::bs_icon("info-circle"), " Hakk\u0131nda"),
    card(card_body(
      tags$h4("Bitki Numuneleme \u0130stasyonlar\u0131 Platformu"),
      tags$p(class = "help-note", style = "font-size:.92rem;",
        "Arazi numuneleme istasyonlar\u0131n\u0131 kaydetmek, haritada g\u00f6rselle\u015ftirmek ve ",
        "tarihsel olarak izlemek i\u00e7in geli\u015ftirilmi\u015ftir."),
      tags$ul(class = "help-note", style = "font-size:.9rem;",
        tags$li("Konum: GPS d\u00fc\u011fmesi veya haritaya dokunma ile nokta se\u00e7ilir; enlem/boylam g\u00f6r\u00fcn\u00fcr."),
        tags$li("Kay\u0131t ekleme: se\u00e7ili konuma; t\u00fcr ad\u0131 yaz\u0131ld\u0131k\u00e7a \u00f6neri gelir, yeni t\u00fcr manuel girilebilir."),
        tags$li("Silme: Veri Tablosu'nda sat\u0131r se\u00e7ip 'Se\u00e7ili kayd\u0131 sil'."),
        tags$li("Ge\u00e7mi\u015f: son 20 s\u00fcr\u00fcm saklan\u0131r; istenilen s\u00fcr\u00fcme d\u00f6n\u00fclebilir veya indirilebilir."),
        tags$li("Lejant: ba\u015fl\u0131\u011fa t\u0131klayarak a\u00e7\u0131l\u0131p kapan\u0131r.")),
      tags$hr(),
      tags$p(class = "help-note",
             tags$span(class = "badge-soft", "PhysIQ Dynamic"),
             " \u00b7 R Shiny + leaflet"))) 
  ),

  nav_spacer(),
  nav_item(tags$a(href = "https://connect.posit.cloud/cm-yesilkanat", target = "_blank",
                  bsicons::bs_icon("cloud-fill"), " Posit Connect"))
)

# ---------------------------------------------------------------------
# 3. SERVER
# ---------------------------------------------------------------------
server <- function(input, output, session) {

  rv <- reactiveValues(history = NULL, active = 1L, pt = NULL, fit = TRUE)

  # -- geçmiş başlat --
  init_history <- function() {
    h <- NULL
    if (file.exists(history_path))
      h <- tryCatch(readRDS(history_path), error = function(e) NULL)
    if (is.null(h) || !length(h)) {
      h <- list(list(time = Sys.time(), note = "Ba\u015flang\u0131\u00e7 veri seti",
                     data = load_default()))
    }
    rv$history <- h; rv$active <- length(h)
  }
  init_history()

  persist <- function() {
    tryCatch(saveRDS(rv$history, history_path), error = function(e) NULL)
  }

  export_df <- function(d) {
    data.frame("ID" = d$ID, "ID2" = d$ID2, "Tarih" = d$Tarih, "T\u00fcr" = d$Tur,
      "Bilimsel" = d$Bilimsel, "T\u00fcrk\u00e7e" = d$Turkce,
      "Enlem (DD - N)" = d$Enlem, "Boylam (DD - E)" = d$Boylam, check.names = FALSE)
  }

  # Her sürümü ayrı Excel olarak yaz; klasörü son 20 dosyaya indir
  write_snapshot <- function(v) {
    safe <- gsub("[^A-Za-z0-9]+", "-", substr(v$note %||% "surum", 1, 22))
    fn <- file.path(history_dir, sprintf("%s_%s.xlsx",
      format(v$time, "%Y%m%d_%H%M%S"), safe))
    tryCatch(writexl::write_xlsx(export_df(v$data), fn), error = function(e) NULL)
    fs <- sort(list.files(history_dir, pattern = "\\.xlsx$", full.names = TRUE))
    if (length(fs) > MAX_VERSIONS) unlink(head(fs, length(fs) - MAX_VERSIONS))
  }

  commit <- function(new_df, note, fit = FALSE) {
    v <- list(time = Sys.time(), note = note, data = new_df)
    rv$history <- c(rv$history, list(v))
    if (length(rv$history) > MAX_VERSIONS)
      rv$history <- tail(rv$history, MAX_VERSIONS)
    rv$active <- length(rv$history)
    rv$fit <- fit
    persist()
    write_snapshot(v)
  }

  current_data <- reactive({ rv$history[[rv$active]]$data })

  # tür seçimini güncelle
  observeEvent(current_data(), {
    sp <- sort(unique(current_data()$Bilimsel))
    updatePickerInput(session, "secili_turler", choices = sp, selected = sp)
  }, ignoreNULL = FALSE)

  # ---- filtre ----
  filtreli <- reactive({
    d <- current_data()
    if (!is.null(input$secili_turler))
      d[d$Bilimsel %in% input$secili_turler, , drop = FALSE] else d[0, , drop = FALSE]
  })
  haritalik <- reactive({
    d <- filtreli()
    d[is.finite(d$Enlem) & is.finite(d$Boylam) &
      d$Enlem >= -90 & d$Enlem <= 90 & d$Boylam >= -180 & d$Boylam <= 180, , drop = FALSE]
  })

  observeEvent(input$secili_turler, { rv$fit <- TRUE }, ignoreNULL = FALSE)

  # ---- değer kutuları ----
  output$vb_istasyon <- renderText(formatC(nrow(haritalik()), format = "d", big.mark = "."))
  output$vb_tur      <- renderText(length(unique(filtreli()$Bilimsel)))
  output$vb_toplam   <- renderText(formatC(nrow(current_data()), format = "d", big.mark = "."))
  output$vb_guncel   <- renderText(format(rv$history[[rv$active]]$time, "%d.%m.%Y %H:%M"))

  # ---- harita ilk çizim ----
  output$harita <- renderLeaflet({
    leaflet() |>
      addProviderTiles(providers$CartoDB.Positron, group = "Sade") |>
      addProviderTiles(providers$Esri.WorldImagery, group = "Uydu") |>
      addProviderTiles(providers$OpenTopoMap, group = "Topografik") |>
      addLayersControl(baseGroups = c("Sade", "Uydu", "Topografik"),
        options = layersControlOptions(collapsed = TRUE)) |>
      addScaleBar(position = "bottomleft") |>
      addFullscreenControl(position = "topleft") |>
      addControlGPS(options = gpsOptions(position = "topleft", activate = FALSE,
        autoCenter = TRUE, setView = TRUE, maxZoom = 15)) |>
      addEasyButton(easyButton(icon = "fa-plus", title = "Bu konuma kay\u0131t ekle",
        position = "topleft",
        onClick = JS("function(btn,map){ Shiny.setInputValue('ekle_harita', Date.now()); }"))) |>
      setView(lng = 41.75, lat = 41.22, zoom = 9)
  })

  # ---- işaretçiler ----
  observe({
    d <- haritalik(); pal <- build_palette(current_data()$Bilimsel)
    proxy <- leafletProxy("harita") |> clearGroup("ist") |> removeControl("uyari")
    if (nrow(d) == 0) {
      proxy |> addControl("Se\u00e7ili t\u00fcr i\u00e7in g\u00f6sterilecek kay\u0131t yok.",
        position = "topright", layerId = "uyari"); return()
    }
    d$renk <- pal[d$Bilimsel]
    popup <- sprintf(
      "<div style='font-size:13px;line-height:1.5'>
        <b style='color:%s'>%s</b>%s<br/>
        <span style='color:#555'>Kod:</span> %s &nbsp;
        <span style='color:#555'>Tarih:</span> %s<br/>
        <span style='color:#555'>\u0130stasyon No:</span> %s &nbsp;
        <span style='color:#555'>Kay\u0131t:</span> %s<br/>
        <span style='color:#555'>Enlem:</span> %.5f &nbsp;
        <span style='color:#555'>Boylam:</span> %.5f</div>",
      d$renk, d$Bilimsel,
      ifelse(nzchar(d$Turkce), paste0(" <span style='color:#777'>(", d$Turkce, ")</span>"), ""),
      d$Tur, ifelse(is.na(d$Tarih), "\u2014", d$Tarih), d$ID2, d$ID, d$Enlem, d$Boylam)
    popup <- enc2utf8(popup)
    lopt <- if (isTRUE(input$etiket))
      labelOptions(noHide = TRUE, textsize = "10px", direction = "top",
                   opacity = 0.85, style = list(padding = "1px 5px")) else labelOptions()
    proxy <- proxy |> addCircleMarkers(data = d, lng = ~Boylam, lat = ~Enlem,
      radius = 7, weight = 1.5, color = "#fff", fillColor = ~renk, fillOpacity = .9,
      popup = popup, label = ~ID2, labelOptions = lopt, group = "ist",
      clusterOptions = if (isTRUE(input$kumelele)) markerClusterOptions() else NULL)
    if (isTRUE(isolate(rv$fit))) {
      if (nrow(d) == 1) proxy |> setView(d$Boylam[1], d$Enlem[1], 13)
      else proxy |> fitBounds(min(d$Boylam), min(d$Enlem), max(d$Boylam), max(d$Enlem))
      rv$fit <- FALSE
    }
  })

  # ---- açılır/kapanır lejant (tür kümesi değişince yenilenir) ----
  observe({
    proxy <- leafletProxy("harita") |> removeControl("lejant")
    if (!isTRUE(input$lejant_goster)) return()
    sp <- sort(unique(haritalik()$Bilimsel)); if (!length(sp)) return()
    pal <- build_palette(current_data()$Bilimsel)
    rows <- paste0(sprintf(
      "<div class='lej-row'><span class='dot' style='background:%s'></span>%s</div>",
      pal[sp], sp), collapse = "")
    html <- sprintf(
      "<details class='lejant'><summary>T\u00fcr Lejant\u0131 (%d)</summary><div class='lej-body'>%s</div></details>",
      length(sp), rows)
    proxy |> addControl(HTML(enc2utf8(html)), position = "bottomright", layerId = "lejant")
  })

  # ---- seçili konum (GPS / tıklama) ----
  set_pt <- function(lat, lng) {
    if (is.null(lat) || is.null(lng) || is.na(lat) || is.na(lng)) return()
    rv$pt <- list(lat = as.numeric(lat), lng = as.numeric(lng))
  }
  observeEvent(input$harita_click, { set_pt(input$harita_click$lat, input$harita_click$lng) })
  observeEvent(input$harita_gps_located, {
    loc <- input$harita_gps_located
    lat <- loc$coordinates$lat %||% loc$latlng$lat %||% loc$lat
    lng <- loc$coordinates$lng %||% loc$latlng$lng %||% loc$lng
    set_pt(lat, lng)
  })
  observeEvent(input$konum_bul, { session$sendCustomMessage("getloc", list()) })
  observeEvent(input$konum_gps, {
    set_pt(input$konum_gps$lat, input$konum_gps$lng)
    if (!is.null(rv$pt))
      leafletProxy("harita") |> setView(rv$pt$lng, rv$pt$lat, zoom = 15)
  })
  observeEvent(input$konum_gps_err, {
    showNotification(paste0("Konum al\u0131namad\u0131: ", sub(" #.*$", "", input$konum_gps_err),
      ". Haritaya dokunarak da nokta se\u00e7ebilirsiniz."), type = "warning", duration = 6)
  })

  observe({
    proxy <- leafletProxy("harita") |> clearGroup("secili")
    if (is.null(rv$pt)) return()
    proxy |> addAwesomeMarkers(lng = rv$pt$lng, lat = rv$pt$lat, group = "secili",
      icon = awesomeIcons(icon = "plus", library = "fa",
                          markerColor = "red", iconColor = "#ffffff"),
      label = "Konum", options = markerOptions(zIndexOffset = 1000))
  })

  output$konum_kutu <- renderUI({
    if (is.null(rv$pt))
      div(class = "coord-box", "Se\u00e7ili konum yok")
    else
      div(class = "coord-box",
          sprintf("Enlem:  %.5f", rv$pt$lat), tags$br(),
          sprintf("Boylam: %.5f", rv$pt$lng))
  })

  # ---- yeni kayıt ekleme modalı ----
  acik_ekle_modal <- function() {
    d <- current_data()
    sp  <- sort(unique(c(SPECIES_REF$Bilimsel, d$Bilimsel)))
    kod <- sort(unique(c(SPECIES_REF$Kod, d$Tur[d$Tur != "\u2014"])))
    trk <- sort(unique(c(SPECIES_REF$Turkce, d$Turkce[nzchar(d$Turkce)])))
    p <- isolate(rv$pt)
    if (is.null(p)) {
      la <- if (nrow(d)) mean(d$Enlem, na.rm = TRUE) else 41.22
      lo <- if (nrow(d)) mean(d$Boylam, na.rm = TRUE) else 41.75
    } else { la <- p$lat; lo <- p$lng }
    yeni_id  <- next_seq(d$ID)
    yeni_id2 <- next_seq(d$ID2)
    showModal(modalDialog(
      title = tagList(icon("seedling"), " Bu konuma yeni kay\u0131t ekle"),
      easyClose = TRUE,
      div(class = "help-note", style = "margin-bottom:.6rem;",
          "Enlem/boylam se\u00e7ili konumdan otomatik gelir. T\u00fcr ad\u0131n\u0131 yazmaya ba\u015flay\u0131n; ",
          "listede yoksa yeni t\u00fcr olarak ekleyebilirsiniz."),
      div(class = "coord-box", style = "margin-bottom:.6rem;",
          sprintf("Otomatik kay\u0131t no (ID): %d", yeni_id)),
      fluidRow(
        column(6, numericInput("m_lat", "Enlem (DD - N)", round(la, 6), step = .00001)),
        column(6, numericInput("m_lon", "Boylam (DD - E)", round(lo, 6), step = .00001))),
      fluidRow(
        column(6, textInput("m_id2", "\u0130stasyon No (ID2)", as.character(yeni_id2))),
        column(6, textInput("m_tarih", "Tarih", format(Sys.Date(), "%Y-%m-%d")))),
      selectizeInput("m_bilimsel", "Bilimsel ad", choices = sp, selected = character(0),
        options = list(create = TRUE, placeholder = "\u00f6rn. Cistus creticus",
          createOnBlur = TRUE)),
      fluidRow(
        column(6, selectizeInput("m_tur", "T\u00fcr (kod)", choices = kod, selected = character(0),
          options = list(create = TRUE, placeholder = "\u00f6rn. CCR"))),
        column(6, selectizeInput("m_turkce", "T\u00fcrk\u00e7e ad", choices = trk,
          selected = character(0),
          options = list(create = TRUE, placeholder = "\u00f6rn. Pembe laden")))),
      footer = tagList(modalButton("\u0130ptal"),
        actionButton("m_kaydet", "Ekle", icon = icon("check"), class = "btn-success"))
    ))
  }
  observeEvent(input$kayit_ekle_btn, acik_ekle_modal())
  observeEvent(input$ekle_harita,   acik_ekle_modal())

  # Referans listesi + mevcut verinin birleşik üçlü tablosu (REF önce → öncelikli)
  combo_table <- reactive({
    d <- current_data()
    dd <- data.frame(Bilimsel = d$Bilimsel, Turkce = d$Turkce, Kod = d$Tur,
                     stringsAsFactors = FALSE)
    dd <- dd[nzchar(dd$Bilimsel) & dd$Bilimsel != "Belirtilmemi\u015f", , drop = FALSE]
    ct <- rbind(SPECIES_REF, dd)
    ct[!duplicated(paste(ct$Bilimsel, ct$Turkce, ct$Kod)), , drop = FALSE]
  })

  # Üç yönlü otomatik tamamlama — döngüsüz ve zamanlamadan bağımsız (deterministik).
  # Bir alan programatik doldurulurken "beklenen değer" not edilir; o alanın
  # geri dönen yankı olayı tanınıp tüketilir, böylece karşılıklı tetikleme olmaz.
  expv <- reactiveValues()
  norm <- function(x) enc2utf8(trimws(as.character(x %||% "")))
  set_field <- function(id, val) {
    if (is.null(val) || !length(val) || is.na(val) || !nzchar(val)) return()
    expv[[id]] <- norm(val)
    updateSelectizeInput(session, id, selected = val)
  }
  is_echo <- function(id, cur) {
    if (!is.null(expv[[id]]) && identical(norm(cur), expv[[id]])) {
      expv[[id]] <- NULL; return(TRUE)
    }
    FALSE
  }
  lookup_row <- function(field, val) {
    ct <- combo_table()
    i <- which(enc2utf8(ct[[field]]) == enc2utf8(norm(val)))[1]
    if (is.na(i)) NULL else ct[i, , drop = FALSE]
  }

  # Bilimsel ad seçilince → Türkçe + kod
  observeEvent(input$m_bilimsel, {
    b <- input$m_bilimsel %||% ""
    if (is_echo("m_bilimsel", b) || !nzchar(b)) return()
    r <- lookup_row("Bilimsel", b); if (is.null(r)) return()
    set_field("m_turkce", r$Turkce); set_field("m_tur", r$Kod)
  }, ignoreInit = TRUE)

  # Türkçe ad seçilince → Bilimsel + kod
  observeEvent(input$m_turkce, {
    tk <- input$m_turkce %||% ""
    if (is_echo("m_turkce", tk) || !nzchar(tk)) return()
    r <- lookup_row("Turkce", tk); if (is.null(r)) return()
    set_field("m_bilimsel", r$Bilimsel); set_field("m_tur", r$Kod)
  }, ignoreInit = TRUE)

  # Kod (Tür) seçilince → Bilimsel + Türkçe
  observeEvent(input$m_tur, {
    kd <- input$m_tur %||% ""
    if (is_echo("m_tur", kd) || !nzchar(kd)) return()
    r <- lookup_row("Kod", kd); if (is.null(r)) return()
    set_field("m_bilimsel", r$Bilimsel); set_field("m_turkce", r$Turkce)
  }, ignoreInit = TRUE)

  observeEvent(input$m_kaydet, {
    bil <- trimws(input$m_bilimsel %||% "")
    if (!nzchar(bil)) { showNotification("Bilimsel ad zorunludur.", type = "error"); return() }
    if (!is.finite(input$m_lat) || !is.finite(input$m_lon)) {
      showNotification("Ge\u00e7erli enlem/boylam giriniz.", type = "error"); return() }
    d <- current_data()
    id_new  <- as.character(next_seq(d$ID))
    id2_new <- { x <- trimws(input$m_id2 %||% ""); if (nzchar(x)) x else as.character(next_seq(d$ID2)) }
    yeni <- data.frame(
      ID = id_new, ID2 = id2_new,
      Tarih = parse_tarih(trimws(input$m_tarih %||% format(Sys.Date(), "%Y-%m-%d"))),
      Tur = { k <- trimws(input$m_tur %||% ""); if (nzchar(k)) k else "\u2014" },
      Bilimsel = bil, Turkce = trimws(input$m_turkce %||% ""),
      Enlem = as.numeric(input$m_lat), Boylam = as.numeric(input$m_lon),
      stringsAsFactors = FALSE)
    commit(rbind(d, yeni), sprintf("Yeni kay\u0131t: ID %s / \u0130st %s (%s)", id_new, id2_new, bil), fit = FALSE)
    removeModal()
    showNotification(sprintf("Kay\u0131t eklendi (ID %s, \u0130stasyon %s).", id_new, id2_new), type = "message")
  })

  # ---- dosya yükleme ----
  observeEvent(input$dosya, {
    req(input$dosya)
    res <- tryCatch(parse_plant_data(as.data.frame(
      readxl::read_excel(input$dosya$datapath, col_types = "text"))),
      error = function(e) e)
    if (inherits(res, "error")) {
      showNotification(paste("Y\u00fckleme hatas\u0131:", conditionMessage(res)),
                       type = "error", duration = 8); return() }
    if (identical(input$yukleme_modu, "append")) {
      commit(dplyr::bind_rows(current_data(), res),
             sprintf("Excel y\u00fcklendi (+%d, ekleme)", nrow(res)), fit = TRUE)
    } else {
      commit(res, sprintf("Excel y\u00fcklendi (%d, de\u011fi\u015ftirme)", nrow(res)), fit = TRUE)
    }
    showNotification(sprintf("%d kay\u0131t y\u00fcklendi.", nrow(res)), type = "message")
  })

  observeEvent(input$sifirla, {
    commit(load_default(), "Ba\u015flang\u0131\u00e7 verisine d\u00f6n\u00fcld\u00fc", fit = TRUE)
    showNotification("Ba\u015flang\u0131\u00e7 verisi y\u00fcklendi.", type = "message")
  })

  output$sablon <- downloadHandler(
    filename = function() "bitki_istasyon_sablonu.xlsx",
    content = function(file) {
      writexl::write_xlsx(data.frame(
        "ID" = c(1, 2),
        "ID2" = c("BN-1", "BN-2"),
        "Tarih" = c("2024", format(Sys.Date(), "%Y-%m-%d")),
        "T\u00fcr" = c("CCR", "URT"),
        "Bilimsel" = c("Cistus creticus", "Urtica dioica"),
        "T\u00fcrk\u00e7e" = c("Pembe laden", "Is\u0131rgan otu"),
        "Enlem (DD - N)" = c(41.1775, 41.1627),
        "Boylam (DD - E)" = c(41.8494, 41.7939), check.names = FALSE), file) }
  )

  # ---- veri tablosu ----
  tablo_df <- reactive({
    d <- current_data(); if (!nrow(d)) return(d)
    d[order(date_key(d$Tarih), decreasing = TRUE), , drop = FALSE]
  })
  output$tablo <- DT::renderDT({
    d <- tablo_df()
    idn <- suppressWarnings(as.integer(d$ID))
    id_col <- if (length(d$ID) && any(is.na(idn) & nzchar(d$ID))) d$ID else idn
    disp <- data.frame(ID = id_col, "ID2" = d$ID2,
      Tarih = d$Tarih, "T\u00fcr" = d$Tur,
      "Bilimsel ad" = d$Bilimsel, "T\u00fcrk\u00e7e" = d$Turkce,
      Enlem = d$Enlem, Boylam = d$Boylam, check.names = FALSE)
    DT::datatable(disp, rownames = FALSE, selection = "single", filter = "top",
      options = list(pageLength = 15, scrollX = TRUE,
        order = list(),                       # tarih-azalan (en güncel üstte) sıra korunur
        columnDefs = list(list(className = "dt-right", targets = 0)),
        language = list(search = "Ara:", info = "_TOTAL_ kayd\u0131n _START_-_END_ aras\u0131",
          paginate = list(previous = "\u00d6nceki", `next` = "Sonraki"))),
      class = "compact stripe hover") |> DT::formatRound(c("Enlem", "Boylam"), 5)
  })

  observeEvent(input$sil_btn, {
    sel <- input$tablo_rows_selected
    if (is.null(sel) || !length(sel)) {
      showNotification("\u00d6nce silinecek sat\u0131r\u0131 se\u00e7in.", type = "warning"); return() }
    d <- tablo_df(); kid <- d$ID[sel]
    showModal(modalDialog(title = "Kayd\u0131 sil",
      sprintf("'%s' kayd\u0131 silinsin mi? Bu i\u015flem yeni bir s\u00fcr\u00fcm olu\u015fturur.", kid),
      footer = tagList(modalButton("Vazge\u00e7"),
        actionButton("sil_onay", "Sil", class = "btn-danger")), easyClose = TRUE))
  })
  observeEvent(input$sil_onay, {
    sel <- isolate(input$tablo_rows_selected); d0 <- isolate(tablo_df())
    kid <- d0$ID[sel]; full <- current_data()
    commit(full[full$ID != kid, , drop = FALSE], sprintf("Silindi: %s", kid), fit = FALSE)
    removeModal(); showNotification(sprintf("%s silindi.", kid), type = "message")
  })

  output$indir_excel <- downloadHandler(
    filename = function() paste0("bitki_aktif_", Sys.Date(), ".xlsx"),
    content = function(file) writexl::write_xlsx(export_df(current_data()), file))

  # ---- özet ----
  ozet <- reactive({
    d <- filtreli(); if (!nrow(d)) return(data.frame(Bilimsel = character(), n = integer()))
    d |> dplyr::group_by(Bilimsel) |> dplyr::summarise(n = dplyr::n(), .groups = "drop") |>
      dplyr::arrange(dplyr::desc(n))
  })
  output$ozet_tablo <- DT::renderDT(
    DT::datatable(ozet(), rownames = FALSE,
      colnames = c("Bilimsel ad", "\u0130stasyon say\u0131s\u0131"),
      options = list(pageLength = 20, dom = "t"), class = "compact stripe hover"))
  output$ozet_grafik <- renderPlot({
    df <- ozet(); if (!nrow(df)) { plot.new(); text(.5, .5, "Veri yok"); return() }
    df <- df[order(df$n), ]; pal <- build_palette(current_data()$Bilimsel)
    op <- par(mar = c(4.5, 12, 1, 2), las = 1); on.exit(par(op))
    bp <- barplot(df$n, horiz = TRUE, names.arg = df$Bilimsel, col = pal[df$Bilimsel],
      border = NA, xlab = "\u0130stasyon say\u0131s\u0131", cex.names = .95, font = 3)
    text(df$n, bp, labels = df$n, pos = 4, cex = .85, xpd = TRUE)
  })

  # ---- sürüm geçmişi ----
  surum_df <- reactive({
    h <- rv$history; n <- length(h)
    data.frame(
      "S\u00fcr\u00fcm" = rev(seq_len(n)),
      "Zaman" = rev(vapply(h, function(v) format(v$time, "%d.%m.%Y %H:%M:%S"), "")),
      "Kay\u0131t" = rev(vapply(h, function(v) nrow(v$data), integer(1))),
      "\u0130\u015flem" = rev(vapply(h, function(v) v$note %||% "", "")),
      "Aktif" = rev(ifelse(seq_len(n) == rv$active, "\u25cf", "")),
      check.names = FALSE)
  })
  output$surum_tablo <- DT::renderDT(
    DT::datatable(surum_df(), rownames = FALSE, selection = "single",
      options = list(pageLength = 20, dom = "tp",
        language = list(paginate = list(previous = "\u00d6nceki", `next` = "Sonraki"))),
      class = "compact stripe hover"))

  selected_version_index <- function() {
    sel <- input$surum_tablo_rows_selected; if (is.null(sel) || !length(sel)) return(NA_integer_)
    rev(seq_len(length(rv$history)))[sel]
  }
  observeEvent(input$geri_don, {
    idx <- selected_version_index()
    if (is.na(idx)) { showNotification("\u00d6nce bir s\u00fcr\u00fcm se\u00e7in.", type = "warning"); return() }
    src <- rv$history[[idx]]
    commit(src$data, sprintf("S\u00fcr\u00fcm %d'e d\u00f6n\u00fcld\u00fc (%s)", idx,
      format(src$time, "%d.%m %H:%M")), fit = TRUE)
    showNotification("Se\u00e7ili s\u00fcr\u00fcme d\u00f6n\u00fcld\u00fc.", type = "message")
  })
  output$indir_surum <- downloadHandler(
    filename = function() paste0("bitki_surum_", Sys.Date(), ".xlsx"),
    content = function(file) {
      idx <- selected_version_index(); if (is.na(idx)) idx <- rv$active
      writexl::write_xlsx(export_df(rv$history[[idx]]$data), file) })
  output$indir_gecmis <- downloadHandler(
    filename = function() paste0("bitki_tum_gecmis_", Sys.Date(), ".xlsx"),
    content = function(file) {
      h <- rv$history; n <- length(h)
      sheets <- setNames(
        lapply(seq_len(n), function(i) export_df(h[[i]]$data)),
        vapply(seq_len(n), function(i)
          substr(sprintf("S%02d_%s", i, format(h[[i]]$time, "%m%d_%H%M")), 1, 31), ""))
      writexl::write_xlsx(sheets, file) })

  # İlk açılışta başlangıç sürümünün Excel anlık görüntüsünü oluştur
  isolate({
    if (!length(list.files(history_dir, pattern = "\\.xlsx$")))
      write_snapshot(rv$history[[rv$active]])
  })
}

shinyApp(ui, server)
