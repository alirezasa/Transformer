// ساخت گزارش Word (report/report.docx) از روی results/metrics.json و تصاویر figures/
// اجرا از ریشه‌ی پروژه:  node report/make_report.js
const fs = require("fs");
const path = require("path");
const {
  Document, Packer, Paragraph, TextRun, ImageRun, Table, TableRow, TableCell, AlignmentType,
  HeadingLevel, WidthType, ShadingType, BorderStyle, PageBreak, Footer, PageNumber, TableOfContents,
  LevelFormat,
} = require("docx");

const ROOT = path.resolve(__dirname, "..");
const M = JSON.parse(fs.readFileSync(path.join(ROOT, "results/metrics.json"), "utf8"));
const A = JSON.parse(fs.readFileSync(path.join(__dirname, "analysis.json"), "utf8"));
const FIG = (f) => path.join(ROOT, "figures", f);

const FA_FONT = "B Nazanin";
const EN_FONT = "Times New Roman";
const SIZE = 26; // half-points -> 13pt
const CONTENT_W = 9638; // A4 (11906) - 2 * 1134 margins

// ---------- helpers ----------
const faDigits = (s) => String(s).replace(/[0-9]/g, (d) => "۰۱۲۳۴۵۶۷۸۹"[d]);
const pct = (x, d = 2) => (100 * x).toFixed(d) + "%";

// متن را به تکه‌های فارسی و لاتین تقسیم می‌کند تا جهت هر تکه درست باشد
function runs(text, opts = {}) {
  if (!/[؀-ۿ]/.test(text)) text = "`" + text + "`"; // متن کاملاً لاتین -> یک تکه‌ی چپ‌به‌راست
  const parts = text.split(/(`[^`]+`|[\[(]?[A-Za-z][A-Za-z0-9_\[\]\-\/.,()+=×:'√ᵀ→·]*(?:\s+[\[(]?[A-Za-z0-9_\[\]\-\/.,()+=×:'√ᵀ→·]+)*)/);
  const out = [];
  for (const part of parts) {
    if (!part) continue;
    const code = part.startsWith("`");
    const t = code ? part.slice(1, -1) : part;
    const latin = code || /^[\[(]?[A-Za-z]/.test(t);
    out.push(new TextRun({
      text: t, bold: opts.bold, italics: opts.italics, size: opts.size || SIZE, color: opts.color,
      rightToLeft: !latin,
      font: latin ? EN_FONT : { ascii: EN_FONT, hAnsi: EN_FONT, cs: FA_FONT },
      ...(latin ? {} : { boldComplexScript: opts.bold, sizeComplexScript: opts.size || SIZE }),
    }));
  }
  return out;
}

const P = (text, opts = {}) => new Paragraph({
  bidirectional: true, alignment: opts.align || AlignmentType.JUSTIFIED,
  spacing: { after: opts.after ?? 120, line: 300 }, children: runs(text, opts),
});
const H = (text, level = HeadingLevel.HEADING_1) => new Paragraph({
  heading: level, bidirectional: true, alignment: AlignmentType.RIGHT,
  spacing: { before: 240, after: 120 }, children: runs(text, { bold: true, size: level === HeadingLevel.HEADING_1 ? 32 : 28, color: "1F3864" }),
});
const BULLET = (text) => new Paragraph({
  bidirectional: true, alignment: AlignmentType.JUSTIFIED, numbering: { reference: "bullets", level: 0 },
  spacing: { after: 60, line: 300 }, children: runs(text),
});
const EQ = (text) => new Paragraph({
  alignment: AlignmentType.CENTER, spacing: { before: 60, after: 120 },
  children: [new TextRun({ text, font: "Cambria Math", size: 24, italics: true })],
});
const CAPTION = (text) => P(text, { align: AlignmentType.CENTER, size: 22, italics: true, color: "404040", after: 200 });

function pngSize(file) {
  const b = fs.readFileSync(file);
  return { w: b.readUInt32BE(16), h: b.readUInt32BE(20) };
}
function IMG(file, widthPx = 620) {
  const { w, h } = pngSize(file);
  return new Paragraph({
    alignment: AlignmentType.CENTER, spacing: { before: 120, after: 60 },
    children: [new ImageRun({ type: "png", data: fs.readFileSync(file), transformation: { width: widthPx, height: Math.round(widthPx * h / w) } })],
  });
}

const border = { style: BorderStyle.SINGLE, size: 4, color: "8EA9DB" };
const borders = { top: border, bottom: border, left: border, right: border };
function TABLE(header, rows, widths) {
  const total = widths.reduce((a, b) => a + b, 0);
  const cell = (text, i, head) => new TableCell({
    width: { size: widths[i], type: WidthType.DXA }, borders,
    shading: head ? { fill: "D9E2F3", type: ShadingType.CLEAR, color: "auto" } : undefined,
    margins: { top: 60, bottom: 60, left: 100, right: 100 },
    children: [new Paragraph({ bidirectional: /[\u0600-\u06FF]/.test(String(text)), alignment: AlignmentType.CENTER, children: runs(String(text), { bold: head, size: 22 }) })],
  });
  return new Table({
    width: { size: total, type: WidthType.DXA }, columnWidths: widths, visuallyRightToLeft: true,
    alignment: AlignmentType.CENTER,
    rows: [
      new TableRow({ tableHeader: true, children: header.map((t, i) => cell(t, i, true)) }),
      ...rows.map((r) => new TableRow({ children: r.map((t, i) => cell(t, i, false)) })),
    ],
  });
}
const SPACER = () => new Paragraph({ spacing: { after: 120 }, children: [] });
const PAGEBREAK = () => new Paragraph({ children: [new PageBreak()] });

// ---------- data ----------
const C = M.config;
const H_ = M.history;
const CLS = ["World", "Sports", "Business", "Sci/Tech"];
const CR = M.classification_report;
const CM = M.confusion_matrix;

// ---------- content ----------
const body = [];

// صفحه‌ی عنوان
body.push(
  new Paragraph({ spacing: { before: 2400 }, children: [] }),
  P("به نام خدا", { align: AlignmentType.CENTER, bold: true, size: 30 }),
  new Paragraph({ spacing: { before: 600 }, children: [] }),
  P("تمرین ۲ — یادگیری عمیق پیشرفته", { align: AlignmentType.CENTER, bold: true, size: 36, color: "1F3864" }),
  P("طراحی و پیاده‌سازی ترنسفورمر از صفر برای طبقه‌بندی متن", { align: AlignmentType.CENTER, bold: true, size: 32 }),
  P("Transformer Encoder from Scratch — AG News", { align: AlignmentType.CENTER, size: 26, italics: true }),
  new Paragraph({ spacing: { before: 1200 }, children: [] }),
  P("نام و نام خانوادگی: ............................", { align: AlignmentType.CENTER }),
  P("شماره‌ی دانشجویی: ............................", { align: AlignmentType.CENTER }),
  P("استاد درس: ............................", { align: AlignmentType.CENTER }),
  P("تاریخ: ............................", { align: AlignmentType.CENTER }),
  PAGEBREAK(),
);

// فهرست
body.push(
  P("فهرست مطالب", { bold: true, size: 30, align: AlignmentType.RIGHT }),
  new TableOfContents("فهرست مطالب", { hyperlink: true, headingStyleRange: "1-2" }),
  P("(در Word روی فهرست راست‌کلیک کرده و Update Field را بزنید تا شماره صفحات به‌روز شود.)", { size: 20, color: "808080" }),
  PAGEBREAK(),
);

// ۱. مقدمه
body.push(
  H("۱. مقدمه و صورت مسئله"),
  P("هدف این تمرین پیاده‌سازی یک کدگذار ترنسفورمر (Transformer Encoder) از صفر و استفاده از آن برای طبقه‌بندی خبرهای مجموعه‌داده‌ی AG News در چهار کلاس World، Sports، Business و Sci/Tech است. طبق الزامات تمرین، استفاده از torch.nn.TransformerEncoder، torch.nn.MultiheadAttention و مدل‌های آماده‌ی Hugging Face ممنوع است؛ بنابراین همه‌ی لایه‌ها فقط با nn.Linear، nn.LayerNorm، nn.Embedding و عملیات ماتریسی تنسورها (matmul، softmax، transpose و view) ساخته شده‌اند. در انتهای نوت‌بوک نیز با یک assert بررسی شده که هیچ‌یک از ماژول‌های ممنوعه در مدل وجود ندارد."),
  P("ساختار گزارش: ابتدا داده و پیش‌پردازش، سپس معماری و کلاس‌های پیاده‌شده، جدول فراپارامترها، تست ابعاد تنسورها، روند آموزش (منحنی‌های Loss و Accuracy)، ارزیابی روی داده‌ی تست (ماتریس درهم‌ریختگی) و در پایان مصورسازی و تحلیل وزن‌های توجه آمده است."),
);

// ۲. داده
const nTrain = 120000 - C.val_size;
body.push(
  H("۲. داده و پیش‌پردازش"),
  P(`مجموعه‌داده‌ی AG News شامل ۱۲۰٬۰۰۰ نمونه‌ی آموزش و ۷٬۶۰۰ نمونه‌ی تست است که میان چهار کلاس کاملاً متوازن تقسیم شده‌اند (هر کلاس ۳۰٬۰۰۰ نمونه‌ی آموزش و ۱٬۹۰۰ نمونه‌ی تست). ${faDigits(C.val_size.toLocaleString("en-US"))} نمونه‌ی تصادفی از داده‌ی آموزش به‌عنوان مجموعه‌ی اعتبارسنجی (Validation) کنار گذاشته شد تا انتخاب بهترین مدل بدون دیدن داده‌ی تست انجام شود.`),
  TABLE(["مجموعه", "تعداد نمونه", "کاربرد"], [
    ["Train", faDigits(nTrain.toLocaleString("en-US")), "به‌روزرسانی وزن‌ها"],
    ["Validation", faDigits(C.val_size.toLocaleString("en-US")), "انتخاب بهترین اپوک"],
    ["Test", "۷٬۶۰۰", "ارزیابی نهایی"],
  ], [2600, 2600, 3600]),
  SPACER(),
  P("مراحل پیش‌پردازش:", { bold: true }),
  BULLET("متن ورودی هر نمونه از اتصال عنوان و توضیح خبر ساخته می‌شود."),
  BULLET("پاک‌سازی: تبدیل به حروف کوچک و حذف تگ‌ها و موجودیت‌های HTML (مانند #36; که به $ تبدیل می‌شود)."),
  BULLET("توکنایزر ساده‌ی سطح کلمه با عبارت منظم: هر کلمه/عدد یک توکن و هر علامت نگارشی یک توکن جدا."),
  BULLET(`ساخت واژگان فقط از روی داده‌ی آموزش با حداقل تکرار ${faDigits(C.min_freq)} و حداکثر اندازه‌ی ${faDigits(C.vocab_size.toLocaleString("en-US"))} واژه. توکن‌های ویژه [PAD]، [UNK] و [CLS] به ترتیب شناسه‌های ۰، ۱ و ۲ را دارند.`),
  BULLET(`افزودن توکن [CLS] به ابتدای هر جمله و بریدن طول به حداکثر ${faDigits(C.max_len)} توکن. میانگین طول جملات حدود ۴۵ توکن و صدک ۹۵ام حدود ۶۴ توکن است، بنابراین کمتر از ۰٫۲٪ نمونه‌ها بریده می‌شوند.`),
  BULLET("Padding پویا: در هر batch جملات فقط تا طول بلندترین جمله‌ی همان batch پد می‌شوند و یک ماسک بولی به نام key_mask ساخته می‌شود که موقعیت‌های پدینگ را مشخص می‌کند."),
);

// ۳. معماری
body.push(
  H("۳. معماری مدل — بخش اول: کلاس‌های مجزا"),
  P("هر جزء مسئله به‌صورت یک کلاس مجزا (زیرکلاس nn.Module) تعریف شده است. جدول زیر کلاس‌ها، ابعاد ورودی و خروجی و نقش هر کدام را نشان می‌دهد (B اندازه‌ی batch، T طول دنباله، d = d_model، h تعداد سرها):"),
  TABLE(["کلاس", "ورودی → خروجی", "توضیح"], [
    ["TokenEmbedding", "(B,T) → (B,T,d)", "جدول امبدینگ × جذر d_model"],
    ["PositionalEncoding", "(B,T,d) → (B,T,d)", "کدگذاری سینوسی موقعیت + Dropout"],
    ["ScaledDotProductAttention", "(B,h,T,d_k) → (B,h,T,d_v)", "softmax(QKᵀ/√d_k)·V با ماسک پدینگ"],
    ["MultiHeadAttention", "(B,T,d) → (B,T,d)", "نگاشت‌های W_Q، W_K، W_V، W_O و تقسیم به h سر"],
    ["PositionwiseFeedForward", "(B,T,d) → (B,T,d)", "Linear(d→d_ff) → GELU → Linear(d_ff→d)"],
    ["EncoderLayer", "(B,T,d) → (B,T,d)", "MHA و FFN با اتصال باقیمانده و LayerNorm"],
    ["TransformerEncoder", "(B,T) → (B,T,d)", "Embedding + PE + پشته‌ی N لایه"],
    ["TransformerClassifier", "(B,T) → (B,4)", "Encoder + Pooling + Linear نهایی"],
  ], [2900, 2700, 4000]),
  SPACER(),
  H("۳-۱. کدگذاری موقعیت", HeadingLevel.HEADING_2),
  P("چون توجه به ترتیب توکن‌ها حساس نیست، اطلاعات موقعیت با کدگذاری سینوسی مقاله‌ی اصلی به امبدینگ‌ها اضافه می‌شود:"),
  EQ("PE(pos, 2i) = sin(pos / 10000^(2i/d_model)),   PE(pos, 2i+1) = cos(pos / 10000^(2i/d_model))"),
  H("۳-۲. توجه ضرب نقطه‌ای مقیاس‌شده و چندسر", HeadingLevel.HEADING_2),
  EQ("Attention(Q, K, V) = softmax( Q·Kᵀ / √d_k + Mask ) · V"),
  P("در کلاس ScaledDotProductAttention امتیازهای مربوط به کلیدهای پدینگ قبل از softmax با −∞ جایگزین می‌شوند تا وزن توجه به آن‌ها دقیقاً صفر شود. در MultiHeadAttention ابتدا ورودی با سه لایه‌ی خطی به Q، K و V نگاشته می‌شود، سپس با view و transpose به ابعاد `B×h×T×d_k` درمی‌آید، توجه برای هر سر به‌صورت موازی محاسبه می‌شود و خروجی سرها دوباره به ابعاد `B×T×d` ادغام و از W_O عبور داده می‌شود:"),
  EQ("MultiHead(X) = Concat(head_1, …, head_h) · W_O,   head_i = Attention(X·W_Q^i, X·W_K^i, X·W_V^i)"),
  P("ماتریس توجه هر لایه در خروجی برگردانده و نگه داشته می‌شود تا در بخش مصورسازی استفاده شود."),
  H("۳-۳. لایه‌ی کدگذار", HeadingLevel.HEADING_2),
  P("هر لایه‌ی کدگذار مانند مقاله‌ی اصلی (Post-LN) از دو زیرلایه با اتصال باقیمانده تشکیل شده است:"),
  EQ("X' = LayerNorm( X + Dropout(MHA(X)) ),    Y = LayerNorm( X' + Dropout(FFN(X')) )"),
  H("۳-۴. بخش دوم: سر طبقه‌بندی (Classification Head)", HeadingLevel.HEADING_2),
  P("هر دو روش خواسته‌شده پیاده‌سازی شده و با پارامتر pooling قابل انتخاب است:"),
  BULLET("روش ۱ — توکن [CLS]: بردار خروجی موقعیت صفر (توکن [CLS]) در لایه‌ی آخر به‌عنوان نماینده‌ی کل جمله استفاده می‌شود."),
  BULLET("روش ۲ — Mean Pooling: میانگین خروجی همه‌ی توکن‌های غیرپدینگ با استفاده از ماسک (m_t برای توکن واقعی ۱ و برای پدینگ ۰ است):"),
  EQ("h = Σ_t m_t · x_t  /  Σ_t m_t"),
  P(`در این گزارش روش «${M.config.pooling === "cls" ? "[CLS]" : "Mean Pooling"}» استفاده شده است. بردار حاصل پس از Dropout از یک لایه‌ی خطی نهایی عبور می‌کند که بعد ${faDigits(C.d_model)} را به ۴ کلاس نگاشت می‌کند و خروجی آن (logits) با تابع زیان Cross-Entropy مقایسه می‌شود.`),
  P(`تعداد کل پارامترهای قابل آموزش ${faDigits(M.n_params.toLocaleString("en-US"))} است که ${faDigits((M.n_params - M.n_params_no_emb).toLocaleString("en-US"))} پارامتر آن مربوط به جدول امبدینگ و فقط ${faDigits(M.n_params_no_emb.toLocaleString("en-US"))} پارامتر مربوط به لایه‌های ترنسفورمر و سر طبقه‌بندی است.`),
);

// ۴. فراپارامترها
body.push(
  H("۴. جدول فراپارامترها"),
  TABLE(["فراپارامتر", "مقدار", "توضیح"], [
    ["d_model", C.d_model, "بعد امبدینگ و مدل"],
    ["h (num_heads)", C.num_heads, "تعداد سرهای توجه"],
    ["d_k = d_v", C.d_k, "d_model / h"],
    ["d_ff", C.d_ff, "بعد لایه‌ی میانی FFN"],
    ["N (num_layers)", C.num_layers, "تعداد لایه‌های کدگذار"],
    ["Pooling", C.pooling === "cls" ? "[CLS]" : "Mean", "استخراج بردار جمله"],
    ["Dropout", C.dropout, "در امبدینگ، توجه، FFN و سر طبقه‌بندی"],
    ["max_len", C.max_len, "حداکثر طول دنباله (با [CLS])"],
    ["Vocabulary size", M.vocab_size, "اندازه‌ی واژگان"],
    ["Batch size", C.batch_size, "اندازه‌ی دسته"],
    ["Epochs", C.epochs, "تعداد اپوک"],
    ["Optimizer", "AdamW", `weight decay = ${C.weight_decay}`],
    ["Learning rate", C.lr, "بیشینه‌ی نرخ یادگیری"],
    ["LR schedule", "Warmup + Cosine", `warmup = ${C.warmup_ratio * 100}% گام‌ها`],
    ["Gradient clipping", C.grad_clip, "نُرم بیشینه‌ی گرادیان"],
    ["Loss", "Cross-Entropy", ""],
    ["Device", M.device.toUpperCase(), `PyTorch ${M.torch_version.split("+")[0]}`],
  ], [3000, 2600, 4000]),
);

// ۵. تست ابعاد
body.push(
  H("۵. تست ابعاد تنسورها در لایه‌ها"),
  P("در نوت‌بوک یک batch واقعی (۴ جمله) از تک‌تک اجزای مدل عبور داده شده و ابعاد خروجی هر مرحله چاپ و با assert بررسی شده است. خلاصه‌ی خروجی (برای B=4 و طول T برابر بلندترین جمله‌ی batch) در جدول زیر آمده است:"),
  TABLE(["مرحله", "ابعاد خروجی"], [
    ["input_ids / key_mask", "(B, T)"],
    ["TokenEmbedding / + PositionalEncoding", "(B, T, 128)"],
    ["Q, K, V (پس از تقسیم به سرها)", "(B, 4, T, 32)"],
    ["خروجی ScaledDotProductAttention", "(B, 4, T, 32)"],
    ["وزن‌های توجه", "(B, 4, T, T)"],
    ["ادغام سرها / خروجی MultiHeadAttention", "(B, T, 128)"],
    ["لایه‌ی میانی FFN (fc1)", "(B, T, 256)"],
    ["خروجی FFN / EncoderLayer 1 و 2", "(B, T, 128)"],
    ["بردار جمله پس از Pooling", "(B, 128)"],
    ["logits", "(B, 4)"],
  ], [5200, 3600]),
  SPACER(),
  P("علاوه بر ابعاد، دو بررسی صحت نیز انجام شده است: (۱) جمع هر سطر ماتریس توجه برابر ۱ است و (۲) وزن توجه روی همه‌ی توکن‌های پدینگ دقیقاً صفر است؛ یعنی ماسک به‌درستی کار می‌کند."),
);

// ۶. آموزش
const epochRows = H_.train_loss.map((_, i) => [
  faDigits(i + 1), H_.train_loss[i].toFixed(4), pct(H_.train_acc[i]), H_.val_loss[i].toFixed(4), pct(H_.val_acc[i]),
]);
body.push(
  H("۶. روند آموزش و منحنی‌های همگرایی"),
  P(`مدل به مدت ${faDigits(C.epochs)} اپوک با اندازه‌ی دسته‌ی ${faDigits(C.batch_size)} روی CPU آموزش داده شد (زمان کل حدود ${faDigits(Math.round(H_.epoch_time.reduce((a, b) => a + b, 0) / 60))} دقیقه). جدول و نمودار زیر Loss و Accuracy را روی داده‌های آموزش و اعتبارسنجی در پایان هر اپوک نشان می‌دهد.`),
  TABLE(["اپوک", "Train Loss", "Train Acc", "Val Loss", "Val Acc"], epochRows, [1400, 2000, 2000, 2000, 2000]),
  IMG(FIG("training_curves.png"), 620),
  CAPTION("شکل ۱ — منحنی همگرایی Loss (چپ) و Accuracy (راست) برای داده‌های آموزش و اعتبارسنجی"),
  P("تحلیل:", { bold: true }),
  ...A.training.map(BULLET),
);

// ۷. ارزیابی
const perClass = CLS.map((c) => [c, pct(CR[c].precision), pct(CR[c].recall), pct(CR[c]["f1-score"]), faDigits(CR[c].support)]);
perClass.push(["Macro avg", pct(CR["macro avg"].precision), pct(CR["macro avg"].recall), pct(CR["macro avg"]["f1-score"]), "۷۶۰۰"]);
body.push(
  H("۷. ارزیابی روی داده‌ی تست"),
  P(`بهترین مدل (بر اساس دقت اعتبارسنجی = ${pct(M.best_val_acc)}) روی ۷٬۶۰۰ نمونه‌ی تست ارزیابی شد: دقت تست ${pct(M.test_acc)} و Loss تست ${M.test_loss.toFixed(4)}.`, { bold: true }),
  TABLE(["کلاس", "Precision", "Recall", "F1-score", "تعداد"], perClass, [2200, 1900, 1900, 1900, 1600]),
  IMG(FIG("confusion_matrix.png"), 620),
  CAPTION("شکل ۲ — ماتریس درهم‌ریختگی روی داده‌ی تست (چپ: تعداد، راست: نرمال‌شده بر اساس سطر)"),
  P("تحلیل ماتریس درهم‌ریختگی:", { bold: true }),
  ...A.confusion.map(BULLET),
);

// ۸. توجه
const att = M.attention;
const topStr = (k) => att[k].top_cls_tokens.map(([t, w]) => `${t} (${w.toFixed(2)})`).join("، ");
body.push(
  PAGEBREAK(),
  H("۸. بخش چهارم: مصورسازی وزن‌های توجه"),
  P("دو جمله از مجموعه‌ی تست انتخاب شد: یک خبر ورزشی که نام تیم (Yankees) در آن آمده و یک خبر فناوری با واژه‌های تخصصی (Skype، software، Wi-Fi). هر دو جمله کوتاه هستند تا heatmap خوانا باشد و هر دو به‌درستی طبقه‌بندی شده‌اند. برای هر جمله ماتریس توجه T×T هر چهار سر در لایه‌ی آخر (لایه‌ی ۲) رسم شده است؛ سطرها query و ستون‌ها key (توکنی که به آن توجه می‌شود) هستند. چون طبقه‌بندی از روی بردار [CLS] انجام می‌شود، مهم‌ترین سطر، سطر اول (توجه [CLS] به سایر کلمات) است که جداگانه نیز رسم شده است."),
);
const examples = [["sports", "ورزشی", "۳", "۴"], ["scitech", "فناوری", "۵", "۶"]];
for (const [k, fa, f1, f2] of examples) {
  const e = att[k];
  const pred = Object.entries(e.probs).sort((a, b) => b[1] - a[1])[0];
  body.push(
    H(`۸-${k === "sports" ? "۱" : "۲"}. جمله‌ی ${fa}`, HeadingLevel.HEADING_2),
    new Paragraph({ alignment: AlignmentType.LEFT, spacing: { after: 120 }, children: [new TextRun({ text: `"${e.text}"`, font: EN_FONT, size: 22, italics: true })] }),
    P(`پیش‌بینی مدل: ${pred[0]} با احتمال ${pct(pred[1], 1)}. پرتوجه‌ترین توکن‌ها از دید [CLS] (میانگین ۴ سر): ${topStr(k)}.`),
    IMG(FIG(`attention_heads_${k}.png`), 600),
    CAPTION(`شکل ${f1} — نقشه‌ی حرارتی توجه چهار سر لایه‌ی آخر برای جمله‌ی ${fa}`),
    IMG(FIG(`attention_cls_${k}.png`), 620),
    CAPTION(`شکل ${f2} — توجه توکن [CLS] به هر توکن در لایه‌ی آخر (هر سر و میانگین سرها)`),
    P("تحلیل:", { bold: true }),
    ...A[k].map(BULLET),
  );
}
body.push(
  H("۸-۳. جمع‌بندی تحلیل توجه", HeadingLevel.HEADING_2),
  ...A.attention_general.map(BULLET),
);

// ۹. نتیجه‌گیری
body.push(
  H("۹. نتیجه‌گیری"),
  ...A.conclusion.map(BULLET),
  H("مراجع"),
  new Paragraph({ alignment: AlignmentType.LEFT, spacing: { after: 100 }, children: [new TextRun({ text: "[1] A. Vaswani et al., \"Attention Is All You Need\", NeurIPS 2017.", font: EN_FONT, size: 22 })] }),
  new Paragraph({ alignment: AlignmentType.LEFT, spacing: { after: 100 }, children: [new TextRun({ text: "[2] X. Zhang, J. Zhao, Y. LeCun, \"Character-level Convolutional Networks for Text Classification\", NeurIPS 2015 (AG News).", font: EN_FONT, size: 22 })] }),
  new Paragraph({ alignment: AlignmentType.LEFT, spacing: { after: 100 }, children: [new TextRun({ text: "[3] J. Devlin et al., \"BERT: Pre-training of Deep Bidirectional Transformers for Language Understanding\", NAACL 2019 ([CLS] token).", font: EN_FONT, size: 22 })] }),
);

// ---------- document ----------
const doc = new Document({
  creator: "Student",
  title: "Transformer Encoder from Scratch — AG News",
  styles: {
    default: { document: { run: { font: { ascii: EN_FONT, hAnsi: EN_FONT, cs: FA_FONT }, size: SIZE, sizeComplexScript: SIZE } } },
    paragraphStyles: [
      { id: "Heading1", name: "Heading 1", basedOn: "Normal", next: "Normal", quickFormat: true,
        run: { size: 32, bold: true, color: "1F3864", font: { ascii: EN_FONT, hAnsi: EN_FONT, cs: FA_FONT } },
        paragraph: { spacing: { before: 240, after: 120 }, outlineLevel: 0 } },
      { id: "Heading2", name: "Heading 2", basedOn: "Normal", next: "Normal", quickFormat: true,
        run: { size: 28, bold: true, color: "1F3864", font: { ascii: EN_FONT, hAnsi: EN_FONT, cs: FA_FONT } },
        paragraph: { spacing: { before: 200, after: 100 }, outlineLevel: 1 } },
    ],
  },
  numbering: {
    config: [{ reference: "bullets", levels: [{ level: 0, format: LevelFormat.BULLET, text: "•", alignment: AlignmentType.RIGHT,
      style: { paragraph: { indent: { right: 500, hanging: 300 } } } }] }],
  },
  sections: [{
    properties: { page: { size: { width: 11906, height: 16838 }, margin: { top: 1134, bottom: 1134, left: 1134, right: 1134 } } },
    footers: { default: new Footer({ children: [new Paragraph({ alignment: AlignmentType.CENTER, children: [new TextRun({ children: [PageNumber.CURRENT], size: 20 })] })] }) },
    children: body,
  }],
});

const out = path.join(__dirname, "report.docx");
Packer.toBuffer(doc).then((buf) => { fs.writeFileSync(out, buf); console.log("wrote", out); });
