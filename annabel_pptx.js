const pptxgen = require("pptxgenjs");
const React = require("react");
const ReactDOMServer = require("react-dom/server");
const sharp = require("sharp");

// Colors — Berry & Cream adapted for beauty salon
const C = {
  dark:    "2D1B2E", // deep plum (dark slides)
  berry:   "7B3F6E", // berry (accent)
  rose:    "C4718A", // dusty rose (secondary)
  cream:   "FAF6F1", // cream (light bg)
  white:   "FFFFFF",
  muted:   "8A7A85",
  text:    "2D1B2E",
  light:   "F0E8EC",
  gold:    "D4956A", // warm accent
};

const { FaInstagram, FaVk, FaMapMarkerAlt, FaAward, FaUsers, FaCamera, FaRegCalendarCheck, FaHeart, FaStar } = require("react-icons/fa");
const { MdPhotoCamera, MdLocationOn, MdTrendingUp, MdPeople, MdFavorite } = require("react-icons/md");
const { BiMap } = require("react-icons/bi");

async function iconPng(IconComponent, color = "#FFFFFF", size = 256) {
  const svg = ReactDOMServer.renderToStaticMarkup(
    React.createElement(IconComponent, { color, size: String(size) })
  );
  const buf = await sharp(Buffer.from(svg)).png().toBuffer();
  return "image/png;base64," + buf.toString("base64");
}

async function main() {
  const pres = new pptxgen();
  pres.layout = "LAYOUT_16x9";
  pres.title = "Концепция продвижения — Annabel";

  const makeShadow = () => ({ type: "outer", blur: 8, offset: 3, angle: 135, color: "000000", opacity: 0.12 });

  // ─── SLIDE 1: TITLE ─────────────────────────────────────────────────────────
  {
    const s = pres.addSlide();
    s.background = { color: C.dark };

    // Left accent block
    s.addShape(pres.shapes.RECTANGLE, { x: 0, y: 0, w: 0.3, h: 5.625, fill: { color: C.berry }, line: { color: C.berry } });

    // Decorative circle top-right
    s.addShape(pres.shapes.OVAL, { x: 7.5, y: -0.8, w: 3.5, h: 3.5, fill: { color: C.berry, transparency: 80 }, line: { color: C.berry, transparency: 80 } });
    s.addShape(pres.shapes.OVAL, { x: 8.2, y: -0.2, w: 2.2, h: 2.2, fill: { color: C.rose, transparency: 70 }, line: { color: C.rose, transparency: 70 } });

    s.addText("КОНЦЕПЦИЯ ПРОДВИЖЕНИЯ", {
      x: 0.6, y: 0.7, w: 8.5, h: 0.5,
      fontSize: 11, fontFace: "Calibri", color: C.rose,
      charSpacing: 4, bold: true, margin: 0
    });

    s.addText("Салон красоты\nAnnabel", {
      x: 0.6, y: 1.2, w: 7, h: 1.9,
      fontSize: 54, fontFace: "Georgia", color: C.white,
      bold: true, margin: 0
    });

    s.addText("Санкт-Петербург · м. Пионерская · 11 лет на рынке", {
      x: 0.6, y: 3.2, w: 7, h: 0.4,
      fontSize: 13, fontFace: "Calibri", color: C.muted, margin: 0
    });

    // Bottom stat row
    const stats = [
      { val: "5 каналов", label: "продвижения" },
      { val: "10–15", label: "новых клиентов/мес (цель)" },
      { val: "20 000 ₽", label: "бюджет/мес" },
    ];
    stats.forEach((st, i) => {
      const x = 0.6 + i * 3.1;
      s.addShape(pres.shapes.RECTANGLE, { x, y: 4.2, w: 2.7, h: 1.0, fill: { color: C.berry, transparency: 60 }, line: { color: C.berry, transparency: 60 } });
      s.addText(st.val, { x, y: 4.22, w: 2.7, h: 0.45, fontSize: 18, fontFace: "Georgia", color: C.white, bold: true, align: "center", margin: 0 });
      s.addText(st.label, { x, y: 4.65, w: 2.7, h: 0.35, fontSize: 10, fontFace: "Calibri", color: C.light, align: "center", margin: 0 });
    });
  }

  // ─── SLIDE 2: ИСХОДНАЯ ТОЧКА ─────────────────────────────────────────────────
  {
    const s = pres.addSlide();
    s.background = { color: C.cream };

    s.addText("С чего начинаем", {
      x: 0.5, y: 0.35, w: 9, h: 0.6,
      fontSize: 28, fontFace: "Georgia", color: C.text, bold: true, margin: 0
    });
    s.addText("Ситуация на старте — честный взгляд на точку А", {
      x: 0.5, y: 0.95, w: 9, h: 0.35,
      fontSize: 13, fontFace: "Calibri", color: C.muted, margin: 0
    });

    // Cards
    const cards = [
      { icon: "❌", title: "Соцсети", body: "Аккаунты в VK и Instagram есть, но заброшены. Заявок — 0." },
      { icon: "📍", title: "Авито", body: "Работает! 2 заявки в месяц — потенциал для масштабирования." },
      { icon: "🔇", title: "Реклама", body: "Почти не запускали. Только сарафанное радио и редкие посты." },
      { icon: "📋", title: "Запись", body: "Онлайн-записи нет. Принимают по телефону и мессенджерам." },
      { icon: "👥", title: "База клиентов", body: "База есть, но для повторных продаж почти не используется." },
      { icon: "📉", title: "Постоянники", body: "Доля упала с 70% до 40–45%. Удержание нужно восстановить." },
    ];

    cards.forEach((c, i) => {
      const col = i % 3;
      const row = Math.floor(i / 3);
      const x = 0.5 + col * 3.1;
      const y = 1.5 + row * 1.8;
      s.addShape(pres.shapes.RECTANGLE, { x, y, w: 2.85, h: 1.55, fill: { color: C.white }, line: { color: "E8D8E0" }, shadow: makeShadow() });
      s.addText(c.icon + "  " + c.title, { x: x + 0.15, y: y + 0.15, w: 2.55, h: 0.35, fontSize: 13, fontFace: "Calibri", color: C.berry, bold: true, margin: 0 });
      s.addText(c.body, { x: x + 0.15, y: y + 0.5, w: 2.55, h: 0.9, fontSize: 11, fontFace: "Calibri", color: C.text, margin: 0 });
    });
  }

  // ─── SLIDE 3: ПОЗИЦИОНИРОВАНИЕ ───────────────────────────────────────────────
  {
    const s = pres.addSlide();
    s.background = { color: C.dark };

    s.addShape(pres.shapes.RECTANGLE, { x: 0, y: 0, w: 0.3, h: 5.625, fill: { color: C.rose }, line: { color: C.rose } });

    s.addText("Голос бренда", {
      x: 0.6, y: 0.4, w: 9, h: 0.55,
      fontSize: 28, fontFace: "Georgia", color: C.white, bold: true, margin: 0
    });
    s.addText("Как Annabel говорит с клиентами", {
      x: 0.6, y: 0.95, w: 9, h: 0.35,
      fontSize: 13, fontFace: "Calibri", color: C.muted, margin: 0
    });

    // Left: persona
    s.addShape(pres.shapes.RECTANGLE, { x: 0.6, y: 1.45, w: 4.4, h: 3.7, fill: { color: C.berry, transparency: 75 }, line: { color: C.berry, transparency: 75 } });

    s.addText("Образ бренда", { x: 0.75, y: 1.6, w: 4.1, h: 0.4, fontSize: 11, fontFace: "Calibri", color: C.rose, bold: true, charSpacing: 2, margin: 0 });

    const personas = ["💛  Внимательная подруга", "🏡  Добрый сосед", "🍃  Спокойный и тактичный"];
    personas.forEach((p, i) => {
      s.addText(p, { x: 0.75, y: 2.1 + i * 0.65, w: 4.1, h: 0.5, fontSize: 14, fontFace: "Calibri", color: C.white, margin: 0 });
    });

    s.addText("Обращение на «вы»\nТон — тёплый, без пафоса и давления\nСтиль — дружеский эксперт", {
      x: 0.75, y: 4.0, w: 4.1, h: 0.9,
      fontSize: 11, fontFace: "Calibri", color: C.light, margin: 0
    });

    // Right: what NOT to do
    s.addShape(pres.shapes.RECTANGLE, { x: 5.3, y: 1.45, w: 4.2, h: 3.7, fill: { color: C.white, transparency: 5 }, line: { color: "3D2040" } });
    s.addText("В коммуникации — никогда", { x: 5.45, y: 1.6, w: 3.9, h: 0.4, fontSize: 11, fontFace: "Calibri", color: C.rose, bold: true, charSpacing: 2, margin: 0 });

    const nots = ["🚫  Пафос и громкие обещания", "🚫  Панибратство и сленг", "🚫  Давление и навязчивые продажи"];
    nots.forEach((n, i) => {
      s.addText(n, { x: 5.45, y: 2.1 + i * 0.65, w: 3.9, h: 0.5, fontSize: 13, fontFace: "Calibri", color: C.white, margin: 0 });
    });

    s.addShape(pres.shapes.RECTANGLE, { x: 5.45, y: 3.75, w: 3.9, h: 0.06, fill: { color: C.berry, transparency: 50 }, line: { color: C.berry, transparency: 50 } });
    s.addText("✨ Звёздный мастер — Гаяне\n8 лет в салоне — лицо и голос бренда", {
      x: 5.45, y: 3.9, w: 3.9, h: 0.9,
      fontSize: 11, fontFace: "Calibri", color: C.light, margin: 0
    });
  }

  // ─── SLIDE 4: INSTAGRAM ──────────────────────────────────────────────────────
  {
    const s = pres.addSlide();
    s.background = { color: C.cream };

    // Instagram gradient header block
    s.addShape(pres.shapes.RECTANGLE, { x: 0, y: 0, w: 10, h: 1.55, fill: { color: "C13584" }, line: { color: "C13584" } });
    s.addShape(pres.shapes.RECTANGLE, { x: 0, y: 0, w: 10, h: 1.55, fill: { color: "833AB4", transparency: 50 }, line: { color: "833AB4", transparency: 50 } });

    const igIcon = await iconPng(FaInstagram, "#FFFFFF", 256);
    s.addImage({ data: igIcon, x: 0.4, y: 0.3, w: 0.75, h: 0.75 });
    s.addText("Instagram", { x: 1.25, y: 0.3, w: 5, h: 0.45, fontSize: 26, fontFace: "Georgia", color: C.white, bold: true, margin: 0 });
    s.addText("Витрина работ и трансформаций · Органический рост", { x: 1.25, y: 0.78, w: 7, h: 0.35, fontSize: 12, fontFace: "Calibri", color: "F0D0EE", margin: 0 });
    s.addText("🎯 Приоритет 2", { x: 8.0, y: 0.5, w: 1.7, h: 0.4, fontSize: 11, fontFace: "Calibri", color: C.white, bold: true, align: "center", margin: 0 });

    // Content blocks
    const items = [
      { icon: "📸", title: "Reels с трансформациями", body: "До/после стрижки, окрашивания, маникюра. Формат даёт максимальный органический охват в алгоритмах Instagram." },
      { icon: "💎", title: "LPG + косметология", body: "Показываем уникальные процедуры: LPG по маслу с RF-лифтингом и элос-эпиляцию — то, чего нет у конкурентов." },
      { icon: "🌸", title: "Гаяне ведёт Stories", body: "Мастер с 8-летним стажем — живые stories от первого лица: приёмы, советы по уходу, закулисье салона." },
      { icon: "📅", title: "Акция для подписчиков", body: "Скидка на первый визит и бонус за рекомендацию — анонсируем в stories, создаём повод подписаться." },
    ];

    items.forEach((it, i) => {
      const col = i % 2;
      const row = Math.floor(i / 2);
      const x = 0.4 + col * 4.8;
      const y = 1.75 + row * 1.8;
      s.addShape(pres.shapes.RECTANGLE, { x, y, w: 4.5, h: 1.6, fill: { color: C.white }, line: { color: "E8D8E0" }, shadow: makeShadow() });
      s.addText(it.icon + "  " + it.title, { x: x + 0.2, y: y + 0.18, w: 4.1, h: 0.38, fontSize: 13, fontFace: "Calibri", color: C.berry, bold: true, margin: 0 });
      s.addText(it.body, { x: x + 0.2, y: y + 0.55, w: 4.1, h: 0.9, fontSize: 11, fontFace: "Calibri", color: C.text, margin: 0 });
    });
  }

  // ─── SLIDE 5: ВКОНТАКТЕ ──────────────────────────────────────────────────────
  {
    const s = pres.addSlide();
    s.background = { color: C.cream };

    s.addShape(pres.shapes.RECTANGLE, { x: 0, y: 0, w: 10, h: 1.55, fill: { color: "0077FF" }, line: { color: "0077FF" } });

    const vkIcon = await iconPng(FaVk, "#FFFFFF", 256);
    s.addImage({ data: vkIcon, x: 0.4, y: 0.3, w: 0.75, h: 0.75 });
    s.addText("ВКонтакте", { x: 1.25, y: 0.3, w: 5, h: 0.45, fontSize: 26, fontFace: "Georgia", color: C.white, bold: true, margin: 0 });
    s.addText("Основной рабочий канал · Таргет + сообщество", { x: 1.25, y: 0.78, w: 7, h: 0.35, fontSize: 12, fontFace: "Calibri", color: "D0E8FF", margin: 0 });
    s.addText("🎯 Приоритет 1", { x: 8.0, y: 0.5, w: 1.7, h: 0.4, fontSize: 11, fontFace: "Calibri", color: C.white, bold: true, align: "center", margin: 0 });

    // Left column: content strategy
    s.addShape(pres.shapes.RECTANGLE, { x: 0.4, y: 1.75, w: 4.5, h: 3.5, fill: { color: C.white }, line: { color: "E0E8F5" }, shadow: makeShadow() });
    s.addText("Контент-стратегия", { x: 0.6, y: 1.9, w: 4.1, h: 0.4, fontSize: 14, fontFace: "Georgia", color: "0077FF", bold: true, margin: 0 });

    const content = [
      "📷 Работы мастеров — фото до/после",
      "🎬 Reels: трансформации и процедуры",
      "🌟 Посты от Гаяне — советы и экспертиза",
      "🎁 Акции: скидка на первый визит",
      "📣 Свободные окна и спецпредложения",
    ];
    content.forEach((c, i) => {
      s.addText(c, { x: 0.6, y: 2.4 + i * 0.52, w: 4.1, h: 0.42, fontSize: 12, fontFace: "Calibri", color: C.text, margin: 0 });
    });

    // Right column: targeting
    s.addShape(pres.shapes.RECTANGLE, { x: 5.1, y: 1.75, w: 4.5, h: 3.5, fill: { color: "0077FF" }, line: { color: "0077FF" } });
    s.addText("Таргет", { x: 5.3, y: 1.9, w: 4.1, h: 0.4, fontSize: 14, fontFace: "Georgia", color: C.white, bold: true, margin: 0 });

    const targeting = [
      ["Аудитория", "Женщины 25–50 лет"],
      ["Гео", "3–5 км от м. Пионерская"],
      ["Бюджет", "8–10 000 ₽/мес на таргет"],
      ["Оффер", "Скидка на первый визит"],
      ["Ретаргет", "Посетители сообщества"],
    ];
    targeting.forEach((t, i) => {
      s.addText(t[0].toUpperCase(), { x: 5.3, y: 2.42 + i * 0.52, w: 1.5, h: 0.35, fontSize: 9, fontFace: "Calibri", color: "90C4FF", bold: true, charSpacing: 1, margin: 0 });
      s.addText(t[1], { x: 6.85, y: 2.42 + i * 0.52, w: 2.55, h: 0.35, fontSize: 12, fontFace: "Calibri", color: C.white, margin: 0 });
    });
  }

  // ─── SLIDE 6: КАРТЫ (Яндекс + 2ГИС) ─────────────────────────────────────────
  {
    const s = pres.addSlide();
    s.background = { color: C.cream };

    s.addShape(pres.shapes.RECTANGLE, { x: 0, y: 0, w: 10, h: 1.55, fill: { color: "1A1A1A" }, line: { color: "1A1A1A" } });

    const mapIcon = await iconPng(FaMapMarkerAlt, "#FF4433", 256);
    s.addImage({ data: mapIcon, x: 0.4, y: 0.3, w: 0.65, h: 0.75 });
    s.addText("Яндекс Карты + 2ГИС", { x: 1.2, y: 0.3, w: 6, h: 0.45, fontSize: 26, fontFace: "Georgia", color: C.white, bold: true, margin: 0 });
    s.addText("Бесплатный органический трафик от клиентов «рядом»", { x: 1.2, y: 0.78, w: 7, h: 0.35, fontSize: 12, fontFace: "Calibri", color: "AAAAAA", margin: 0 });
    s.addText("🎯 Приоритет 1", { x: 8.0, y: 0.5, w: 1.7, h: 0.4, fontSize: 11, fontFace: "Calibri", color: C.white, bold: true, align: "center", margin: 0 });

    const steps = [
      { num: "01", title: "Заполнить профиль полностью", body: "Фото интерьера, каждого мастера и работ. Актуальный прайс. Описание всех 6 направлений услуг. График работы." },
      { num: "02", title: "Собрать 30+ отзывов", body: "Попросить постоянных клиентов оставить оценку. Минимум 15–20 отзывов снимают сомнения у новых клиентов перед визитом." },
      { num: "03", title: "Подключить онлайн-запись", body: "Кнопка «Записаться» через Яндекс Бизнес. Прямая конверсия без звонка — снижает порог входа для нового клиента." },
      { num: "04", title: "Отвечать на все отзывы", body: "Благодарить за позитивные, корректно отрабатывать негатив. Это сигнал алгоритму и доверие для читателей." },
    ];

    steps.forEach((st, i) => {
      const col = i % 2;
      const row = Math.floor(i / 2);
      const x = 0.4 + col * 4.8;
      const y = 1.75 + row * 1.85;
      s.addShape(pres.shapes.RECTANGLE, { x, y, w: 4.5, h: 1.65, fill: { color: C.white }, line: { color: "DDDDDD" }, shadow: makeShadow() });
      s.addShape(pres.shapes.RECTANGLE, { x, y, w: 0.65, h: 1.65, fill: { color: "1A1A1A" }, line: { color: "1A1A1A" } });
      s.addText(st.num, { x, y: y + 0.55, w: 0.65, h: 0.55, fontSize: 16, fontFace: "Georgia", color: "FF4433", bold: true, align: "center", margin: 0 });
      s.addText(st.title, { x: x + 0.75, y: y + 0.15, w: 3.6, h: 0.45, fontSize: 12, fontFace: "Calibri", color: C.text, bold: true, margin: 0 });
      s.addText(st.body, { x: x + 0.75, y: y + 0.6, w: 3.6, h: 0.85, fontSize: 10.5, fontFace: "Calibri", color: C.muted, margin: 0 });
    });
  }

  // ─── SLIDE 7: АВИТО ───────────────────────────────────────────────────────────
  {
    const s = pres.addSlide();
    s.background = { color: C.cream };

    s.addShape(pres.shapes.RECTANGLE, { x: 0, y: 0, w: 10, h: 1.55, fill: { color: "00A046" }, line: { color: "00A046" } });
    s.addText("Авито", { x: 0.5, y: 0.3, w: 5, h: 0.45, fontSize: 26, fontFace: "Georgia", color: C.white, bold: true, margin: 0 });
    s.addText("Уже работает — масштабируем с 2 до 10–15 заявок в месяц", { x: 0.5, y: 0.78, w: 7.5, h: 0.35, fontSize: 12, fontFace: "Calibri", color: "C0EDD4", margin: 0 });
    s.addText("🎯 Масштабировать", { x: 7.7, y: 0.5, w: 2.0, h: 0.4, fontSize: 11, fontFace: "Calibri", color: C.white, bold: true, align: "center", margin: 0 });

    // Current vs target
    s.addShape(pres.shapes.RECTANGLE, { x: 0.4, y: 1.75, w: 2.2, h: 2.0, fill: { color: "F5F5F5" }, line: { color: "DDDDDD" } });
    s.addText("СЕЙЧАС", { x: 0.4, y: 1.85, w: 2.2, h: 0.35, fontSize: 10, fontFace: "Calibri", color: C.muted, bold: true, charSpacing: 2, align: "center", margin: 0 });
    s.addText("2", { x: 0.4, y: 2.2, w: 2.2, h: 0.85, fontSize: 64, fontFace: "Georgia", color: "CCCCCC", bold: true, align: "center", margin: 0 });
    s.addText("заявки/мес", { x: 0.4, y: 3.1, w: 2.2, h: 0.4, fontSize: 11, fontFace: "Calibri", color: C.muted, align: "center", margin: 0 });

    s.addText("→", { x: 2.7, y: 2.2, w: 0.6, h: 0.85, fontSize: 32, fontFace: "Georgia", color: "00A046", align: "center", margin: 0 });

    s.addShape(pres.shapes.RECTANGLE, { x: 3.4, y: 1.75, w: 2.2, h: 2.0, fill: { color: "00A046" }, line: { color: "00A046" } });
    s.addText("ЦЕЛЬ", { x: 3.4, y: 1.85, w: 2.2, h: 0.35, fontSize: 10, fontFace: "Calibri", color: "90EDB8", bold: true, charSpacing: 2, align: "center", margin: 0 });
    s.addText("15", { x: 3.4, y: 2.2, w: 2.2, h: 0.85, fontSize: 64, fontFace: "Georgia", color: C.white, bold: true, align: "center", margin: 0 });
    s.addText("заявок/мес", { x: 3.4, y: 3.1, w: 2.2, h: 0.4, fontSize: 11, fontFace: "Calibri", color: "C0EDD4", align: "center", margin: 0 });

    // Right: how
    s.addShape(pres.shapes.RECTANGLE, { x: 5.9, y: 1.75, w: 3.7, h: 2.0, fill: { color: C.white }, line: { color: "CCCCCC" }, shadow: makeShadow() });
    s.addText("Как достичь цели", { x: 6.1, y: 1.9, w: 3.3, h: 0.4, fontSize: 13, fontFace: "Georgia", color: "00A046", bold: true, margin: 0 });
    const avito = [
      "Отдельное объявление для каждой услуги",
      "Платное поднятие топ-объявлений (5–7к ₽/мес)",
      "Фото работ и интерьера в каждом объявлении",
      "Быстрые ответы — конверсия падает за 1 час",
    ];
    avito.forEach((a, i) => {
      s.addText("✓  " + a, { x: 6.1, y: 2.4 + i * 0.3, w: 3.3, h: 0.25, fontSize: 11, fontFace: "Calibri", color: C.text, margin: 0 });
    });

    // Bottom note
    s.addShape(pres.shapes.RECTANGLE, { x: 0.4, y: 4.05, w: 9.2, h: 0.65, fill: { color: "E8F7EE" }, line: { color: "B0DEC0" } });
    s.addText("💡  Авито даёт самый дешёвый лид в бьюти-нише в СПб. Уже есть доверие площадки — нужно просто расширить присутствие.", {
      x: 0.6, y: 4.12, w: 8.8, h: 0.5,
      fontSize: 11, fontFace: "Calibri", color: "1A6B35", margin: 0
    });
  }

  // ─── SLIDE 8: TELEGRAM ────────────────────────────────────────────────────────
  {
    const s = pres.addSlide();
    s.background = { color: C.cream };

    s.addShape(pres.shapes.RECTANGLE, { x: 0, y: 0, w: 10, h: 1.55, fill: { color: "229ED9" }, line: { color: "229ED9" } });
    s.addText("Telegram-канал", { x: 0.5, y: 0.3, w: 6, h: 0.45, fontSize: 26, fontFace: "Georgia", color: C.white, bold: true, margin: 0 });
    s.addText("Удержание базы · Повторные визиты · Свободные окна", { x: 0.5, y: 0.78, w: 7.5, h: 0.35, fontSize: 12, fontFace: "Calibri", color: "C0E8F8", margin: 0 });
    s.addText("🎯 Приоритет 3", { x: 8.0, y: 0.5, w: 1.7, h: 0.4, fontSize: 11, fontFace: "Calibri", color: C.white, bold: true, align: "center", margin: 0 });

    s.addShape(pres.shapes.RECTANGLE, { x: 0.4, y: 1.7, w: 4.5, h: 3.55, fill: { color: C.white }, line: { color: "C5E5F5" }, shadow: makeShadow() });
    s.addText("Для кого", { x: 0.6, y: 1.85, w: 4.1, h: 0.4, fontSize: 14, fontFace: "Georgia", color: "229ED9", bold: true, margin: 0 });
    s.addText("Канал не для привлечения новых —\nтолько для работы с существующей базой клиентов.\n\nПостоянники подпишутся сами, если дать им повод: эксклюзивные акции, приоритетная запись, советы Гаяне.", {
      x: 0.6, y: 2.35, w: 4.1, h: 2.7,
      fontSize: 12, fontFace: "Calibri", color: C.text, margin: 0
    });

    s.addShape(pres.shapes.RECTANGLE, { x: 5.1, y: 1.7, w: 4.5, h: 3.55, fill: { color: "229ED9" }, line: { color: "229ED9" } });
    s.addText("Контент канала", { x: 5.3, y: 1.85, w: 4.1, h: 0.4, fontSize: 14, fontFace: "Georgia", color: C.white, bold: true, margin: 0 });
    const tg = [
      "🗓  Свободные окна на этой неделе",
      "🎁  Закрытые акции только для подписчиков",
      "💬  Советы Гаяне по уходу за волосами",
      "📸  Работы мастеров — новинки и тренды",
      "⭐  Отзывы клиентов и истории трансформаций",
    ];
    tg.forEach((t, i) => {
      s.addText(t, { x: 5.3, y: 2.4 + i * 0.52, w: 4.1, h: 0.42, fontSize: 12, fontFace: "Calibri", color: C.white, margin: 0 });
    });
  }

  // ─── SLIDE 9: ROADMAP ────────────────────────────────────────────────────────
  {
    const s = pres.addSlide();
    s.background = { color: C.dark };
    s.addShape(pres.shapes.RECTANGLE, { x: 0, y: 0, w: 0.3, h: 5.625, fill: { color: C.gold }, line: { color: C.gold } });

    s.addText("Роадмап на 3 месяца", {
      x: 0.6, y: 0.35, w: 9, h: 0.55,
      fontSize: 28, fontFace: "Georgia", color: C.white, bold: true, margin: 0
    });
    s.addText("Что делаем и когда — чтобы выйти на 10–15 новых клиентов в месяц", {
      x: 0.6, y: 0.92, w: 9, h: 0.35,
      fontSize: 13, fontFace: "Calibri", color: C.muted, margin: 0
    });

    const months = [
      {
        num: "1", title: "Фундамент",
        color: C.berry,
        items: ["Заполнить Яндекс Карты и 2ГИС полностью", "Оживить VK: 10–15 постов с работами мастеров", "Расширить объявления на Авито (по каждой услуге)", "Запустить сбор отзывов у постоянных клиентов", "Подключить онлайн-запись"],
      },
      {
        num: "2", title: "Первый трафик",
        color: C.rose,
        items: ["Запустить таргет VK (8–10к ₽/мес)", "Оффер: скидка на первый визит", "Гео: 3–5 км от м. Пионерская", "Instagram: первые Reels от Гаяне", "Поднять топ-объявления на Авито"],
      },
      {
        num: "3", title: "Оптимизация",
        color: C.gold,
        items: ["Анализ: что даёт заявки, масштабируем", "Запуск Telegram-канала для базы клиентов", "Ретаргет VK на посетителей сообщества", "Бонус за рекомендации — системно", "Замер: 10–15 новых клиентов в месяц"],
      },
    ];

    months.forEach((m, i) => {
      const x = 0.5 + i * 3.15;
      s.addShape(pres.shapes.RECTANGLE, { x, y: 1.5, w: 2.95, h: 3.8, fill: { color: m.color, transparency: 80 }, line: { color: m.color, transparency: 60 } });
      s.addShape(pres.shapes.RECTANGLE, { x, y: 1.5, w: 2.95, h: 0.65, fill: { color: m.color }, line: { color: m.color } });
      s.addText("Месяц " + m.num + " — " + m.title, { x: x + 0.15, y: 1.58, w: 2.65, h: 0.5, fontSize: 13, fontFace: "Georgia", color: C.white, bold: true, margin: 0 });
      m.items.forEach((it, j) => {
        s.addText("· " + it, { x: x + 0.15, y: 2.28 + j * 0.56, w: 2.65, h: 0.5, fontSize: 10.5, fontFace: "Calibri", color: C.light, margin: 0 });
      });
    });
  }

  // ─── SLIDE 10: KPI ────────────────────────────────────────────────────────────
  {
    const s = pres.addSlide();
    s.background = { color: C.cream };

    s.addText("KPI через 2 месяца", {
      x: 0.5, y: 0.35, w: 9, h: 0.55,
      fontSize: 28, fontFace: "Georgia", color: C.text, bold: true, margin: 0
    });
    s.addText("Целевые показатели, которые мы считаем успехом", {
      x: 0.5, y: 0.92, w: 9, h: 0.35,
      fontSize: 13, fontFace: "Calibri", color: C.muted, margin: 0
    });

    const kpis = [
      { val: "10–15", label: "новых клиентов/мес", color: C.berry },
      { val: "30+", label: "отзывов на Яндекс Картах", color: "0077FF" },
      { val: "500+", label: "подписчиков ВКонтакте", color: "0077FF" },
      { val: "10+", label: "заявок с Авито/мес", color: "00A046" },
      { val: "×5", label: "охват в соцсетях vs старт", color: C.rose },
      { val: "↑", label: "узнаваемость в районе", color: C.gold },
    ];

    kpis.forEach((k, i) => {
      const col = i % 3;
      const row = Math.floor(i / 3);
      const x = 0.5 + col * 3.1;
      const y = 1.5 + row * 2.0;
      s.addShape(pres.shapes.RECTANGLE, { x, y, w: 2.8, h: 1.75, fill: { color: C.white }, line: { color: "E8D8E0" }, shadow: makeShadow() });
      s.addShape(pres.shapes.RECTANGLE, { x, y, w: 2.8, h: 0.1, fill: { color: k.color }, line: { color: k.color } });
      s.addText(k.val, { x, y: y + 0.2, w: 2.8, h: 0.75, fontSize: 42, fontFace: "Georgia", color: k.color, bold: true, align: "center", margin: 0 });
      s.addText(k.label, { x: x + 0.15, y: y + 1.0, w: 2.5, h: 0.55, fontSize: 12, fontFace: "Calibri", color: C.muted, align: "center", margin: 0 });
    });
  }

  // ─── SLIDE 11: ФИНАЛ ─────────────────────────────────────────────────────────
  {
    const s = pres.addSlide();
    s.background = { color: C.dark };

    s.addShape(pres.shapes.OVAL, { x: -1, y: -1, w: 5, h: 5, fill: { color: C.berry, transparency: 85 }, line: { color: C.berry, transparency: 85 } });
    s.addShape(pres.shapes.OVAL, { x: 7, y: 2.5, w: 4, h: 4, fill: { color: C.rose, transparency: 85 }, line: { color: C.rose, transparency: 85 } });

    s.addText("Annabel — 11 лет доверия.\nПора, чтобы об этом узнал весь район.", {
      x: 0.8, y: 1.2, w: 8.4, h: 1.8,
      fontSize: 30, fontFace: "Georgia", color: C.white, bold: true, align: "center", margin: 0
    });

    s.addText("Следующий шаг — согласовать план и начать с фундамента:\nЯндекс Карты, отзывы, онлайн-запись.", {
      x: 1.5, y: 3.1, w: 7, h: 0.9,
      fontSize: 14, fontFace: "Calibri", color: C.muted, align: "center", margin: 0
    });

    s.addShape(pres.shapes.RECTANGLE, { x: 3.5, y: 4.2, w: 3.0, h: 0.7, fill: { color: C.berry }, line: { color: C.berry }, shadow: makeShadow() });
    s.addText("Начнём?", { x: 3.5, y: 4.28, w: 3.0, h: 0.55, fontSize: 16, fontFace: "Georgia", color: C.white, bold: true, align: "center", margin: 0 });
  }

  await pres.writeFile({ fileName: "/mnt/user-data/outputs/Annabel_Концепция_продвижения.pptx" });
  console.log("Done");
}

main().catch(console.error);
