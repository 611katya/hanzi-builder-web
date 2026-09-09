import React, { useState, useEffect, useMemo, useCallback, useRef } from "react";
import * as XLSX from "xlsx";
import HanziWriter from "hanzi-writer";
import { supabase } from "./supabaseClient.js";

/* ============================================================
   拼字 · GHÉP CHỮ — a Hanzi-building game
   Build complete Chinese characters out of their bushou (部首)
   components. Hints: meaning, pinyin, Sino-Vietnamese reading.
   Includes a manual entry tool to grow the dictionary.
   ============================================================ */

/* ---------- Seed data: 部首 (components) ---------- */
const SEED_BUSHOU = [
  // 1 stroke
  { char: "一", pinyin: "yī", meaning: "one", sv: "nhất", strokes: 1 },
  { char: "丨", pinyin: "gǔn", meaning: "vertical line", sv: "cổn", strokes: 1 },
  { char: "丶", pinyin: "zhǔ", meaning: "dot", sv: "chủ", strokes: 1 },
  { char: "丿", pinyin: "piě", meaning: "slash, left stroke", sv: "phiệt", strokes: 1 },
  { char: "乙", pinyin: "yǐ", meaning: "second; bird", sv: "ất", strokes: 1 },
  { char: "亅", pinyin: "jué", meaning: "hook", sv: "quyết", strokes: 1 },
  // 2 strokes
  { char: "二", pinyin: "èr", meaning: "two", sv: "nhị", strokes: 2 },
  { char: "刀", pinyin: "dāo", meaning: "knife", sv: "đao", strokes: 2 },
  { char: "力", pinyin: "lì", meaning: "strength", sv: "lực", strokes: 2 },
  { char: "又", pinyin: "yòu", meaning: "again; hand", sv: "hựu", strokes: 2 },
  { char: "人", pinyin: "rén", meaning: "person", sv: "nhân", strokes: 2 },
  { char: "亻", pinyin: "rén", meaning: "person (side form)", sv: "nhân", strokes: 2 },
  { char: "讠", pinyin: "yán", meaning: "speech (side form)", sv: "ngôn", strokes: 2 },
  { char: "冫", pinyin: "bīng", meaning: "ice", sv: "băng", strokes: 2 },
  { char: "十", pinyin: "shí", meaning: "ten", sv: "thập", strokes: 2 },
  { char: "卜", pinyin: "bǔ", meaning: "divination", sv: "bốc", strokes: 2 },
  { char: "几", pinyin: "jī", meaning: "small table, stool", sv: "kỷ", strokes: 2 },
  { char: "儿", pinyin: "ér", meaning: "legs; son", sv: "nhi", strokes: 2 },
  { char: "匚", pinyin: "fāng", meaning: "box, container", sv: "phương", strokes: 2 },
  { char: "厂", pinyin: "hǎn", meaning: "cliff, factory", sv: "hán", strokes: 2 },
  { char: "亠", pinyin: "tóu", meaning: "lid, roof top", sv: "đầu", strokes: 2 },
  { char: "入", pinyin: "rù", meaning: "to enter", sv: "nhập", strokes: 2 },
  { char: "八", pinyin: "bā", meaning: "eight", sv: "bát", strokes: 2 },
  { char: "冂", pinyin: "jiōng", meaning: "borders, wide plain", sv: "quynh", strokes: 2 },
  { char: "冖", pinyin: "mì", meaning: "cover, cloth cover", sv: "mịch", strokes: 2 },
  { char: "凵", pinyin: "kǎn", meaning: "open container", sv: "khảm", strokes: 2 },
  { char: "勹", pinyin: "bāo", meaning: "to wrap", sv: "bao", strokes: 2 },
  { char: "匕", pinyin: "bǐ", meaning: "spoon; dagger", sv: "chuỷ", strokes: 2 },
  { char: "匸", pinyin: "xì", meaning: "hiding enclosure", sv: "hệ", strokes: 2 },
  { char: "卩", pinyin: "jié", meaning: "seal, tally", sv: "tiết", strokes: 2 },
  { char: "厶", pinyin: "sī", meaning: "private, personal", sv: "khư", strokes: 2 },
  // 3 strokes
  { char: "口", pinyin: "kǒu", meaning: "mouth", sv: "khẩu", strokes: 3 },
  { char: "土", pinyin: "tǔ", meaning: "earth, soil", sv: "thổ", strokes: 3 },
  { char: "大", pinyin: "dà", meaning: "big", sv: "đại", strokes: 3 },
  { char: "女", pinyin: "nǚ", meaning: "woman", sv: "nữ", strokes: 3 },
  { char: "子", pinyin: "zǐ", meaning: "child", sv: "tử", strokes: 3 },
  { char: "寸", pinyin: "cùn", meaning: "inch", sv: "thốn", strokes: 3 },
  { char: "小", pinyin: "xiǎo", meaning: "small", sv: "tiểu", strokes: 3 },
  { char: "山", pinyin: "shān", meaning: "mountain", sv: "sơn", strokes: 3 },
  { char: "工", pinyin: "gōng", meaning: "work, labor", sv: "công", strokes: 3 },
  { char: "已", pinyin: "yǐ", meaning: "already", sv: "dĩ", strokes: 3 },
  { char: "马", pinyin: "mǎ", meaning: "horse", sv: "mã", strokes: 3 },
  { char: "门", pinyin: "mén", meaning: "gate, door", sv: "môn", strokes: 3 },
  { char: "亡", pinyin: "wáng", meaning: "to flee, perish", sv: "vong", strokes: 3 },
  { char: "也", pinyin: "yě", meaning: "also, too", sv: "dã", strokes: 3 },
  { char: "三", pinyin: "sān", meaning: "three", sv: "tam", strokes: 3 },
  { char: "氵", pinyin: "shuǐ", meaning: "water (side form)", sv: "thủy", strokes: 3 },
  { char: "忄", pinyin: "xīn", meaning: "heart (side form)", sv: "tâm", strokes: 3 },
  { char: "扌", pinyin: "shǒu", meaning: "hand (side form)", sv: "thủ", strokes: 3 },
  { char: "纟", pinyin: "sī", meaning: "silk (side form)", sv: "mịch", strokes: 3 },
  { char: "宀", pinyin: "mián", meaning: "roof", sv: "miên", strokes: 3 },
  { char: "艹", pinyin: "cǎo", meaning: "grass", sv: "thảo", strokes: 3 },
  { char: "阝", pinyin: "fù / yì", meaning: "mound (L) / city (R)", sv: "phụ / ấp", strokes: 3 },
  { char: "尸", pinyin: "shī", meaning: "corpse, body", sv: "thi", strokes: 3 },
  { char: "己", pinyin: "jǐ", meaning: "self", sv: "kỷ", strokes: 3 },
  { char: "巾", pinyin: "jīn", meaning: "cloth, towel", sv: "cân", strokes: 3 },
  { char: "干", pinyin: "gān", meaning: "dry; shield", sv: "can", strokes: 3 },
  { char: "幺", pinyin: "yāo", meaning: "tiny, small", sv: "yêu", strokes: 3 },
  { char: "广", pinyin: "guǎng", meaning: "shelter, lean-to", sv: "nghiễm", strokes: 3 },
  { char: "弓", pinyin: "gōng", meaning: "bow (weapon)", sv: "cung", strokes: 3 },
  { char: "彳", pinyin: "chì", meaning: "step, stride", sv: "xích", strokes: 3 },
  { char: "飞", pinyin: "fēi", meaning: "to fly", sv: "phi", strokes: 3 },
  { char: "辶", pinyin: "chuò", meaning: "walk (side form)", sv: "sước", strokes: 3 },
  { char: "饣", pinyin: "shí", meaning: "food (side form)", sv: "thực", strokes: 3 },
  { char: "囗", pinyin: "wéi", meaning: "enclosure", sv: "vi", strokes: 3 },
  { char: "士", pinyin: "shì", meaning: "scholar, gentleman", sv: "sĩ", strokes: 3 },
  { char: "夂", pinyin: "zhǐ", meaning: "to go, follow", sv: "truy", strokes: 3 },
  { char: "夊", pinyin: "suī", meaning: "to go slowly", sv: "tuy", strokes: 3 },
  { char: "夕", pinyin: "xī", meaning: "evening", sv: "tịch", strokes: 3 },
  { char: "尢", pinyin: "wāng", meaning: "lame, crooked", sv: "uông", strokes: 3 },
  { char: "屮", pinyin: "chè", meaning: "sprout", sv: "triệt", strokes: 3 },
  { char: "巛", pinyin: "chuān", meaning: "river, stream", sv: "xuyên", strokes: 3 },
  { char: "廴", pinyin: "yǐn", meaning: "long stride", sv: "dẫn", strokes: 3 },
  { char: "廾", pinyin: "gǒng", meaning: "clasped hands", sv: "củng", strokes: 3 },
  { char: "弋", pinyin: "yì", meaning: "dart, to shoot", sv: "dặc", strokes: 3 },
  { char: "彐", pinyin: "jì", meaning: "snout, hog's head", sv: "kệ", strokes: 3 },
  { char: "彡", pinyin: "shān", meaning: "bristle, hair lines", sv: "sam", strokes: 3 },
  // 4 strokes
  { char: "日", pinyin: "rì", meaning: "sun, day", sv: "nhật", strokes: 4 },
  { char: "月", pinyin: "yuè", meaning: "moon, month", sv: "nguyệt", strokes: 4 },
  { char: "木", pinyin: "mù", meaning: "tree, wood", sv: "mộc", strokes: 4 },
  { char: "水", pinyin: "shuǐ", meaning: "water", sv: "thủy", strokes: 4 },
  { char: "火", pinyin: "huǒ", meaning: "fire", sv: "hỏa", strokes: 4 },
  { char: "灬", pinyin: "huǒ", meaning: "fire (bottom form)", sv: "hỏa", strokes: 4 },
  { char: "心", pinyin: "xīn", meaning: "heart", sv: "tâm", strokes: 4 },
  { char: "手", pinyin: "shǒu", meaning: "hand", sv: "thủ", strokes: 4 },
  { char: "王", pinyin: "wáng", meaning: "king", sv: "vương", strokes: 4 },
  { char: "牛", pinyin: "niú", meaning: "cattle, ox", sv: "ngưu", strokes: 4 },
  { char: "犬", pinyin: "quǎn", meaning: "dog", sv: "khuyển", strokes: 4 },
  { char: "车", pinyin: "chē", meaning: "cart, vehicle", sv: "xa", strokes: 4 },
  { char: "贝", pinyin: "bèi", meaning: "shell, money", sv: "bối", strokes: 4 },
  { char: "反", pinyin: "fǎn", meaning: "opposite, to turn over", sv: "phản", strokes: 4 },
  { char: "化", pinyin: "huà", meaning: "to change", sv: "hóa", strokes: 4 },
  { char: "长", pinyin: "cháng", meaning: "long", sv: "trường", strokes: 4 },
  { char: "戈", pinyin: "gē", meaning: "spear, dagger-axe", sv: "qua", strokes: 4 },
  { char: "户", pinyin: "hù", meaning: "door, household", sv: "hộ", strokes: 4 },
  { char: "支", pinyin: "zhī", meaning: "branch", sv: "chi", strokes: 4 },
  { char: "攵", pinyin: "pū", meaning: "tap, rap (side form)", sv: "phộc", strokes: 4 },
  { char: "文", pinyin: "wén", meaning: "script, culture", sv: "văn", strokes: 4 },
  { char: "斗", pinyin: "dǒu", meaning: "dipper; to fight", sv: "đấu", strokes: 4 },
  { char: "斤", pinyin: "jīn", meaning: "axe", sv: "cân", strokes: 4 },
  { char: "方", pinyin: "fāng", meaning: "square, direction", sv: "phương", strokes: 4 },
  { char: "无", pinyin: "wú", meaning: "not, without", sv: "vô", strokes: 4 },
  { char: "欠", pinyin: "qiàn", meaning: "to lack; to yawn", sv: "khiếm", strokes: 4 },
  { char: "比", pinyin: "bǐ", meaning: "to compare", sv: "tỷ", strokes: 4 },
  { char: "毛", pinyin: "máo", meaning: "fur, hair", sv: "mao", strokes: 4 },
  { char: "气", pinyin: "qì", meaning: "steam, gas", sv: "khí", strokes: 4 },
  { char: "爪", pinyin: "zhǎo", meaning: "claw", sv: "trảo", strokes: 4 },
  { char: "父", pinyin: "fù", meaning: "father", sv: "phụ", strokes: 4 },
  { char: "片", pinyin: "piàn", meaning: "slice, plank", sv: "phiến", strokes: 4 },
  { char: "牙", pinyin: "yá", meaning: "tooth", sv: "nha", strokes: 4 },
  { char: "礻", pinyin: "shì", meaning: "spirit, altar (side form)", sv: "thị", strokes: 4 },
  { char: "曰", pinyin: "yuē", meaning: "to say", sv: "viết", strokes: 4 },
  { char: "止", pinyin: "zhǐ", meaning: "to stop", sv: "chỉ", strokes: 4 },
  { char: "歹", pinyin: "dǎi", meaning: "death, bad", sv: "ngạt", strokes: 4 },
  { char: "殳", pinyin: "shū", meaning: "weapon, spear-pole", sv: "thù", strokes: 4 },
  { char: "毋", pinyin: "wú", meaning: "do not", sv: "vô", strokes: 4 },
  { char: "氏", pinyin: "shì", meaning: "clan, family name", sv: "thị", strokes: 4 },
  { char: "爻", pinyin: "yáo", meaning: "trigram lines", sv: "hào", strokes: 4 },
  { char: "爿", pinyin: "qiáng", meaning: "half tree trunk", sv: "tường", strokes: 4 },
  { char: "韦", pinyin: "wéi", meaning: "tanned leather", sv: "vi", strokes: 4 },
  // 5 strokes
  { char: "田", pinyin: "tián", meaning: "field", sv: "điền", strokes: 5 },
  { char: "目", pinyin: "mù", meaning: "eye", sv: "mục", strokes: 5 },
  { char: "石", pinyin: "shí", meaning: "stone", sv: "thạch", strokes: 5 },
  { char: "禾", pinyin: "hé", meaning: "grain, cereal plant", sv: "hòa", strokes: 5 },
  { char: "生", pinyin: "shēng", meaning: "life, to be born", sv: "sinh", strokes: 5 },
  { char: "玉", pinyin: "yù", meaning: "jade", sv: "ngọc", strokes: 5 },
  { char: "令", pinyin: "lìng", meaning: "order, command", sv: "lệnh", strokes: 5 },
  { char: "尔", pinyin: "ěr", meaning: "you (archaic)", sv: "nhĩ", strokes: 5 },
  { char: "可", pinyin: "kě", meaning: "may, can", sv: "khả", strokes: 5 },
  { char: "且", pinyin: "qiě", meaning: "moreover", sv: "thả", strokes: 5 },
  { char: "鸟", pinyin: "niǎo", meaning: "bird", sv: "điểu", strokes: 5 },
  { char: "夬", pinyin: "guài", meaning: "resolute (phonetic)", sv: "quải", strokes: 5 },
  { char: "玄", pinyin: "xuán", meaning: "dark, mysterious", sv: "huyền", strokes: 5 },
  { char: "瓜", pinyin: "guā", meaning: "melon", sv: "qua", strokes: 5 },
  { char: "甘", pinyin: "gān", meaning: "sweet", sv: "cam", strokes: 5 },
  { char: "白", pinyin: "bái", meaning: "white", sv: "bạch", strokes: 5 },
  { char: "皮", pinyin: "pí", meaning: "skin, hide", sv: "bì", strokes: 5 },
  { char: "矢", pinyin: "shǐ", meaning: "arrow", sv: "thỉ", strokes: 5 },
  { char: "立", pinyin: "lì", meaning: "to stand", sv: "lập", strokes: 5 },
  { char: "穴", pinyin: "xué", meaning: "cave, hole", sv: "huyệt", strokes: 5 },
  { char: "示", pinyin: "shì", meaning: "spirit, altar", sv: "thị", strokes: 5 },
  { char: "用", pinyin: "yòng", meaning: "to use", sv: "dụng", strokes: 5 },
  { char: "钅", pinyin: "jīn", meaning: "metal (side form)", sv: "kim", strokes: 5 },
  { char: "衤", pinyin: "yī", meaning: "clothing (side form)", sv: "y", strokes: 5 },
  { char: "瓦", pinyin: "wǎ", meaning: "tile", sv: "ngõa", strokes: 5 },
  { char: "疋", pinyin: "pǐ", meaning: "bolt of cloth", sv: "thất", strokes: 5 },
  { char: "疒", pinyin: "nè", meaning: "sickness", sv: "nạch", strokes: 5 },
  { char: "癶", pinyin: "bō", meaning: "footsteps", sv: "bát", strokes: 5 },
  { char: "皿", pinyin: "mǐn", meaning: "dish, vessel", sv: "mãnh", strokes: 5 },
  { char: "矛", pinyin: "máo", meaning: "spear", sv: "mâu", strokes: 5 },
  { char: "禸", pinyin: "róu", meaning: "footprint, track", sv: "nhựu", strokes: 5 },
  { char: "龙", pinyin: "lóng", meaning: "dragon", sv: "long", strokes: 5 },
  // 6 strokes
  { char: "糸", pinyin: "mì", meaning: "silk", sv: "mịch", strokes: 6 },
  { char: "早", pinyin: "zǎo", meaning: "early", sv: "tảo", strokes: 6 },
  { char: "羊", pinyin: "yáng", meaning: "sheep, goat", sv: "dương", strokes: 6 },
  { char: "虫", pinyin: "chóng", meaning: "insect", sv: "trùng", strokes: 6 },
  { char: "页", pinyin: "yè", meaning: "page; head", sv: "hiệt", strokes: 6 },
  { char: "耳", pinyin: "ěr", meaning: "ear", sv: "nhĩ", strokes: 6 },
  { char: "竹", pinyin: "zhú", meaning: "bamboo", sv: "trúc", strokes: 6 },
  { char: "米", pinyin: "mǐ", meaning: "rice", sv: "mễ", strokes: 6 },
  { char: "衣", pinyin: "yī", meaning: "clothing", sv: "y", strokes: 6 },
  { char: "臣", pinyin: "chén", meaning: "minister, official", sv: "thần", strokes: 6 },
  { char: "网", pinyin: "wǎng", meaning: "net", sv: "võng", strokes: 6 },
  { char: "缶", pinyin: "fǒu", meaning: "jar, earthenware", sv: "phẫu", strokes: 6 },
  { char: "羽", pinyin: "yǔ", meaning: "feather", sv: "vũ", strokes: 6 },
  { char: "老", pinyin: "lǎo", meaning: "old", sv: "lão", strokes: 6 },
  { char: "而", pinyin: "ér", meaning: "and yet", sv: "nhi", strokes: 6 },
  { char: "自", pinyin: "zì", meaning: "self; nose", sv: "tự", strokes: 6 },
  { char: "至", pinyin: "zhì", meaning: "to arrive", sv: "chí", strokes: 6 },
  { char: "舌", pinyin: "shé", meaning: "tongue", sv: "thiệt", strokes: 6 },
  { char: "舟", pinyin: "zhōu", meaning: "boat", sv: "chu", strokes: 6 },
  { char: "色", pinyin: "sè", meaning: "color", sv: "sắc", strokes: 6 },
  { char: "血", pinyin: "xuè", meaning: "blood", sv: "huyết", strokes: 6 },
  { char: "行", pinyin: "xíng", meaning: "to go; road", sv: "hành", strokes: 6 },
  { char: "耒", pinyin: "lěi", meaning: "plow", sv: "lỗi", strokes: 6 },
  { char: "聿", pinyin: "yù", meaning: "brush", sv: "duật", strokes: 6 },
  { char: "肉", pinyin: "ròu", meaning: "meat, flesh", sv: "nhục", strokes: 6 },
  { char: "臼", pinyin: "jiù", meaning: "mortar", sv: "cữu", strokes: 6 },
  { char: "舛", pinyin: "chuǎn", meaning: "to oppose; dance steps", sv: "suyễn", strokes: 6 },
  { char: "艮", pinyin: "gèn", meaning: "stopping (trigram)", sv: "cấn", strokes: 6 },
  { char: "虍", pinyin: "hū", meaning: "tiger stripes", sv: "hô", strokes: 6 },
  { char: "西", pinyin: "xī", meaning: "west; to cover", sv: "tây", strokes: 6 },
  { char: "齐", pinyin: "qí", meaning: "even, uniform", sv: "tề", strokes: 6 },
  // 7 strokes
  { char: "言", pinyin: "yán", meaning: "speech, words", sv: "ngôn", strokes: 7 },
  { char: "豕", pinyin: "shǐ", meaning: "pig", sv: "thỉ", strokes: 7 },
  { char: "每", pinyin: "měi", meaning: "each, every", sv: "mỗi", strokes: 7 },
  { char: "吾", pinyin: "wú", meaning: "I, my (archaic)", sv: "ngô", strokes: 7 },
  { char: "兑", pinyin: "duì", meaning: "exchange (phonetic)", sv: "đoái", strokes: 7 },
  { char: "足", pinyin: "zú", meaning: "foot", sv: "túc", strokes: 7 },
  { char: "见", pinyin: "jiàn", meaning: "to see", sv: "kiến", strokes: 7 },
  { char: "角", pinyin: "jiǎo", meaning: "horn", sv: "giác", strokes: 7 },
  { char: "谷", pinyin: "gǔ", meaning: "valley", sv: "cốc", strokes: 7 },
  { char: "豆", pinyin: "dòu", meaning: "bean", sv: "đậu", strokes: 7 },
  { char: "走", pinyin: "zǒu", meaning: "to walk, run", sv: "tẩu", strokes: 7 },
  { char: "身", pinyin: "shēn", meaning: "body", sv: "thân", strokes: 7 },
  { char: "辛", pinyin: "xīn", meaning: "bitter, spicy", sv: "tân", strokes: 7 },
  { char: "辰", pinyin: "chén", meaning: "early morning", sv: "thần", strokes: 7 },
  { char: "酉", pinyin: "yǒu", meaning: "wine", sv: "dậu", strokes: 7 },
  { char: "豸", pinyin: "zhì", meaning: "badger; legless insect", sv: "trãi", strokes: 7 },
  { char: "赤", pinyin: "chì", meaning: "red", sv: "xích", strokes: 7 },
  { char: "里", pinyin: "lǐ", meaning: "village; mile", sv: "lý", strokes: 7 },
  { char: "龟", pinyin: "guī", meaning: "turtle", sv: "quy", strokes: 7 },
  { char: "卤", pinyin: "lǔ", meaning: "salt, brine", sv: "lỗ", strokes: 7 },
  // 8 strokes
  { char: "金", pinyin: "jīn", meaning: "metal, gold", sv: "kim", strokes: 8 },
  { char: "青", pinyin: "qīng", meaning: "blue-green", sv: "thanh", strokes: 8 },
  { char: "雨", pinyin: "yǔ", meaning: "rain", sv: "vũ", strokes: 8 },
  { char: "隹", pinyin: "zhuī", meaning: "short-tailed bird", sv: "chuy", strokes: 8 },
  { char: "阜", pinyin: "fù", meaning: "mound", sv: "phụ", strokes: 8 },
  { char: "非", pinyin: "fēi", meaning: "wrong, not", sv: "phi", strokes: 8 },
  { char: "鱼", pinyin: "yú", meaning: "fish", sv: "ngư", strokes: 8 },
  { char: "釆", pinyin: "biàn", meaning: "to distinguish", sv: "biện", strokes: 8 },
  { char: "隶", pinyin: "lì", meaning: "to reach; servant", sv: "đãi", strokes: 8 },
  { char: "黾", pinyin: "mǐn", meaning: "frog", sv: "mãnh", strokes: 8 },
  { char: "齿", pinyin: "chǐ", meaning: "tooth", sv: "xỉ", strokes: 8 },
  // 9 strokes
  { char: "相", pinyin: "xiāng", meaning: "mutual, each other", sv: "tương", strokes: 9 },
  { char: "音", pinyin: "yīn", meaning: "sound", sv: "âm", strokes: 9 },
  { char: "面", pinyin: "miàn", meaning: "face; surface", sv: "diện", strokes: 9 },
  { char: "香", pinyin: "xiāng", meaning: "fragrant", sv: "hương", strokes: 9 },
  { char: "食", pinyin: "shí", meaning: "food", sv: "thực", strokes: 9 },
  { char: "首", pinyin: "shǒu", meaning: "head", sv: "thủ", strokes: 9 },
  { char: "革", pinyin: "gé", meaning: "leather, hide", sv: "cách", strokes: 9 },
  { char: "韭", pinyin: "jiǔ", meaning: "chives", sv: "cửu", strokes: 9 },
  { char: "骨", pinyin: "gǔ", meaning: "bone", sv: "cốt", strokes: 9 },
  { char: "鬼", pinyin: "guǐ", meaning: "ghost", sv: "quỷ", strokes: 9 },
  // 10 strokes
  { char: "高", pinyin: "gāo", meaning: "tall, high", sv: "cao", strokes: 10 },
  { char: "髟", pinyin: "biāo", meaning: "long hair", sv: "tiêu", strokes: 10 },
  { char: "鬯", pinyin: "chàng", meaning: "sacrificial wine", sv: "sưởng", strokes: 10 },
  { char: "鬥", pinyin: "dòu", meaning: "to fight", sv: "đấu", strokes: 10 },
  { char: "鬲", pinyin: "lì", meaning: "cauldron", sv: "cách", strokes: 10 },
  // 11 strokes
  { char: "鹿", pinyin: "lù", meaning: "deer", sv: "lộc", strokes: 11 },
  { char: "麻", pinyin: "má", meaning: "hemp", sv: "ma", strokes: 11 },
  { char: "黄", pinyin: "huáng", meaning: "yellow", sv: "hoàng", strokes: 11 },
  // 12 strokes
  { char: "黍", pinyin: "shǔ", meaning: "millet", sv: "thử", strokes: 12 },
  { char: "黹", pinyin: "zhǐ", meaning: "embroidery", sv: "chỉ", strokes: 12 },
  // 13 strokes
  { char: "鼎", pinyin: "dǐng", meaning: "tripod, cauldron", sv: "đỉnh", strokes: 13 },
  { char: "鼓", pinyin: "gǔ", meaning: "drum", sv: "cổ", strokes: 13 },
  { char: "鼠", pinyin: "shǔ", meaning: "rat", sv: "thử", strokes: 13 },
  // 14 strokes
  { char: "鼻", pinyin: "bí", meaning: "nose", sv: "tị", strokes: 14 },
  // 17 strokes
  { char: "龠", pinyin: "yuè", meaning: "flute, pipes", sv: "thược", strokes: 17 },
  // added for advanced vocabulary (管辖/削弱/局面/执掌/支持/采纳/竞相/序幕)
  { char: "官", pinyin: "guān", meaning: "official", sv: "quan", strokes: 8 },
  { char: "害", pinyin: "hài", meaning: "harm", sv: "hại", strokes: 10 },
  { char: "肖", pinyin: "xiào", meaning: "resemble", sv: "tiếu", strokes: 7 },
  { char: "刂", pinyin: "dāo", meaning: "knife (side form)", sv: "đao", strokes: 2 },
  { char: "句", pinyin: "jù", meaning: "sentence, phrase", sv: "cú", strokes: 5 },
  { char: "九", pinyin: "jiǔ", meaning: "nine", sv: "cửu", strokes: 2 },
  { char: "尚", pinyin: "shàng", meaning: "still, to esteem", sv: "thượng", strokes: 8 },
  { char: "寺", pinyin: "sì", meaning: "temple", sv: "tự", strokes: 6 },
  { char: "爫", pinyin: "zhǎo", meaning: "claw (top form)", sv: "trảo", strokes: 4 },
  { char: "内", pinyin: "nèi", meaning: "inside", sv: "nội", strokes: 4 },
  { char: "予", pinyin: "yǔ", meaning: "to give", sv: "dư", strokes: 4 },
  { char: "莫", pinyin: "mò", meaning: "do not, none", sv: "mạc", strokes: 10 },
  { char: "丆", pinyin: "hàn", meaning: "slanting stroke (variant top)", sv: "hán", strokes: 2 },
  { char: "由", pinyin: "yóu", meaning: "from, cause", sv: "do", strokes: 5 },
];

/* ---------- Seed data: complete characters ---------- */
const SEED_CHARACTERS = [
  { char: "好", pinyin: "hǎo", meaning: "good, well", sv: "hảo", components: ["女", "子"], lists: ["Cơ bản"] },
  { char: "明", pinyin: "míng", meaning: "bright", sv: "minh", components: ["日", "月"], lists: ["Cơ bản"] },
  { char: "休", pinyin: "xiū", meaning: "to rest", sv: "hưu", components: ["亻", "木"], lists: ["Cơ bản"] },
  { char: "林", pinyin: "lín", meaning: "woods, grove", sv: "lâm", components: ["木", "木"], lists: ["Cơ bản"] },
  { char: "森", pinyin: "sēn", meaning: "forest", sv: "sâm", components: ["木", "木", "木"], lists: ["Cơ bản"] },
  { char: "家", pinyin: "jiā", meaning: "home, family", sv: "gia", components: ["宀", "豕"], lists: ["Cơ bản"] },
  { char: "字", pinyin: "zì", meaning: "character, word", sv: "tự", components: ["宀", "子"], lists: ["Cơ bản"] },
  { char: "安", pinyin: "ān", meaning: "peace, safe", sv: "an", components: ["宀", "女"], lists: ["Cơ bản"] },
  { char: "江", pinyin: "jiāng", meaning: "river", sv: "giang", components: ["氵", "工"], lists: ["Cơ bản"] },
  { char: "河", pinyin: "hé", meaning: "river", sv: "hà", components: ["氵", "可"], lists: ["Cơ bản"] },
  { char: "海", pinyin: "hǎi", meaning: "sea, ocean", sv: "hải", components: ["氵", "每"], lists: ["Cơ bản"] },
  { char: "汉", pinyin: "hàn", meaning: "Han (Chinese)", sv: "hán", components: ["氵", "又"], lists: ["Cơ bản"] },
  { char: "忙", pinyin: "máng", meaning: "busy", sv: "mang", components: ["忄", "亡"], lists: ["Cơ bản"] },
  { char: "快", pinyin: "kuài", meaning: "fast, quick", sv: "khoái", components: ["忄", "夬"], lists: ["Cơ bản"] },
  { char: "情", pinyin: "qíng", meaning: "feeling, emotion", sv: "tình", components: ["忄", "青"], lists: ["Cơ bản"] },
  { char: "想", pinyin: "xiǎng", meaning: "to think, want", sv: "tưởng", components: ["相", "心"], lists: ["Cơ bản"] },
  { char: "思", pinyin: "sī", meaning: "to think", sv: "tư", components: ["田", "心"], lists: ["Cơ bản"] },
  { char: "男", pinyin: "nán", meaning: "male", sv: "nam", components: ["田", "力"], lists: ["Cơ bản"] },
  { char: "语", pinyin: "yǔ", meaning: "language", sv: "ngữ", components: ["讠", "吾"], lists: ["Cơ bản"] },
  { char: "说", pinyin: "shuō", meaning: "to speak, say", sv: "thuyết", components: ["讠", "兑"], lists: ["Cơ bản"] },
  { char: "请", pinyin: "qǐng", meaning: "please; to invite", sv: "thỉnh", components: ["讠", "青"], lists: ["Cơ bản"] },
  { char: "认", pinyin: "rèn", meaning: "to recognize", sv: "nhận", components: ["讠", "人"], lists: ["Cơ bản"] },
  { char: "妈", pinyin: "mā", meaning: "mom", sv: "mã", components: ["女", "马"], lists: ["Cơ bản"] },
  { char: "姐", pinyin: "jiě", meaning: "older sister", sv: "tỷ", components: ["女", "且"], lists: ["Cơ bản"] },
  { char: "星", pinyin: "xīng", meaning: "star", sv: "tinh", components: ["日", "生"], lists: ["Cơ bản"] },
  { char: "晴", pinyin: "qíng", meaning: "sunny, clear sky", sv: "tình", components: ["日", "青"], lists: ["Cơ bản"] },
  { char: "暗", pinyin: "àn", meaning: "dark, dim", sv: "ám", components: ["日", "音"], lists: ["Cơ bản"] },
  { char: "村", pinyin: "cūn", meaning: "village", sv: "thôn", components: ["木", "寸"], lists: ["Cơ bản"] },
  { char: "相", pinyin: "xiāng", meaning: "mutual, each other", sv: "tương", components: ["木", "目"], lists: ["Cơ bản"] },
  { char: "板", pinyin: "bǎn", meaning: "board, plank", sv: "bản", components: ["木", "反"], lists: ["Cơ bản"] },
  { char: "冷", pinyin: "lěng", meaning: "cold", sv: "lãnh", components: ["冫", "令"], lists: ["Cơ bản"] },
  { char: "花", pinyin: "huā", meaning: "flower", sv: "hoa", components: ["艹", "化"], lists: ["Cơ bản"] },
  { char: "草", pinyin: "cǎo", meaning: "grass", sv: "thảo", components: ["艹", "早"], lists: ["Cơ bản"] },
  { char: "你", pinyin: "nǐ", meaning: "you", sv: "nễ", components: ["亻", "尔"], lists: ["Cơ bản"] },
  { char: "他", pinyin: "tā", meaning: "he, him", sv: "tha", components: ["亻", "也"], lists: ["Cơ bản"] },
  { char: "们", pinyin: "men", meaning: "(plural marker)", sv: "môn", components: ["亻", "门"], lists: ["Cơ bản"] },
  // advanced vocabulary, added for 管辖/削弱/局面/执掌/支持/采纳/竞相/序幕
  { char: "管", pinyin: "guǎn", meaning: "tube; to manage, administer", sv: "quản", components: ["竹", "官"], lists: ["Nâng cao"] },
  { char: "辖", pinyin: "xiá", meaning: "to govern, have jurisdiction", sv: "hạt", components: ["车", "害"], lists: ["Nâng cao"] },
  { char: "削", pinyin: "xuē", meaning: "to pare, cut down", sv: "tước", components: ["肖", "刂"], lists: ["Nâng cao"] },
  { char: "弱", pinyin: "ruò", meaning: "weak", sv: "nhược", components: ["弓", "冫", "弓", "冫"], lists: ["Nâng cao"] },
  { char: "局", pinyin: "jú", meaning: "situation; bureau, office", sv: "cục", components: ["尸", "句"], lists: ["Nâng cao"] },
  // note: 面 is technically an atomic Kangxi radical with no standard sub-decomposition —
  // this split is a visual/mnemonic aid for this app only, not classical etymology.
  { char: "面", pinyin: "miàn", meaning: "face; aspect", sv: "diện", components: ["丆", "由"], lists: ["Nâng cao"] },
  // note: 执's classical form is 幸+丸; this uses the common simplified-handwriting
  // visual split (扌+九) as a memorable approximation.
  { char: "执", pinyin: "zhí", meaning: "to hold, carry out", sv: "chấp", components: ["扌", "九"], lists: ["Nâng cao"] },
  { char: "掌", pinyin: "zhǎng", meaning: "palm; to be in charge of", sv: "chưởng", components: ["尚", "手"], lists: ["Nâng cao"] },
  { char: "支", pinyin: "zhī", meaning: "branch; to support", sv: "chi", components: ["十", "又"], lists: ["Nâng cao"] },
  { char: "持", pinyin: "chí", meaning: "to hold, maintain", sv: "trì", components: ["扌", "寺"], lists: ["Nâng cao"] },
  { char: "采", pinyin: "cǎi", meaning: "to pick, gather", sv: "thái", components: ["爫", "木"], lists: ["Nâng cao"] },
  { char: "纳", pinyin: "nà", meaning: "to admit, accept", sv: "nạp", components: ["纟", "内"], lists: ["Nâng cao"] },
  // note: 竞's classical form merges 立+兄; this uses a common simplified visual split.
  { char: "竞", pinyin: "jìng", meaning: "to compete", sv: "cạnh", components: ["立", "儿"], lists: ["Nâng cao"] },
  { char: "序", pinyin: "xù", meaning: "order; preface", sv: "tự", components: ["广", "予"], lists: ["Nâng cao"] },
  { char: "幕", pinyin: "mù", meaning: "curtain, screen", sv: "mạc", components: ["莫", "巾"], lists: ["Nâng cao"] },
];

/* ---------- Seed data: multi-character words ----------
   A word is playable as long as EVERY character in it already exists in
   characterList with its own components — those per-character radicals are
   looked up live via buildCharGroups rather than duplicated here, so words
   automatically become playable/unplayable as the underlying characters
   are added, edited, or removed. */
const SEED_WORDS = [
  { word: "管辖", chars: ["管", "辖"], pinyin: "guǎnxiá", meaning: "to have jurisdiction over, administer", sv: "quản hạt", lists: ["Nâng cao"] },
  { word: "削弱", chars: ["削", "弱"], pinyin: "xuēruò", meaning: "to weaken", sv: "tước nhược", lists: ["Nâng cao"] },
  { word: "局面", chars: ["局", "面"], pinyin: "júmiàn", meaning: "situation, aspect", sv: "cục diện", lists: ["Nâng cao"] },
  { word: "执掌", chars: ["执", "掌"], pinyin: "zhízhǎng", meaning: "to take charge of, wield", sv: "chấp chưởng", lists: ["Nâng cao"] },
  { word: "支持", chars: ["支", "持"], pinyin: "zhīchí", meaning: "to support", sv: "chi trì", lists: ["Nâng cao"] },
  { word: "采纳", chars: ["采", "纳"], pinyin: "cǎinà", meaning: "to adopt, accept (advice)", sv: "thái nạp", lists: ["Nâng cao"] },
  { word: "竞相", chars: ["竞", "相"], pinyin: "jìngxiāng", meaning: "to vie with one another", sv: "cạnh tương", lists: ["Nâng cao"] },
  { word: "序幕", chars: ["序", "幕"], pinyin: "xùmù", meaning: "prologue, prelude", sv: "tự mạc", lists: ["Nâng cao"] },
];

// Looks up each character's own components from characterList. Returns null
// (meaning "not playable yet") if any character is missing or is itself an
// indivisible single component.
function buildCharGroups(chars, characterList) {
  const groups = [];
  for (const ch of chars) {
    const found = characterList.find((c) => c.char === ch);
    if (!found || !Array.isArray(found.components) || found.components.length < 2) return null;
    groups.push({ char: ch, components: found.components });
  }
  return groups;
}

/* ---------- Fonts + design tokens ---------- */
const LOGO_DATA_URI = "data:image/png;base64,iVBORw0KGgoAAAANSUhEUgAAAw8AAAHFCAYAAAC0MMUtAAAAAXNSR0IArs4c6QAAAARnQU1BAACxjwv8YQUAAAAJcEhZcwAAFiUAABYlAUlSJPAAADgSSURBVHhe7d13nxzVmS/w+3ru9XrXu07YeLGNTU4iI5MziIxEzsFEYTImg0gCRAYBwgTLgEBkEDlnbLy+L6Gunub20N3zzMyZ7urpnunvH98PaOpUVYeZ7vrVOec5/+vv//N/KwAAgKkIDwAAQBHhAQAAKCI8AAAARYQHAACgiPAAAAAUER4AAIAiwgMAAFBEeAAAAIoIDwAAQBHhAQAAKCI8AAAARYQHAACgiPAAAAAUER4AAIAiwgMAAFBEeAAAAIoIDwAAQBHhAQAAKCI8AAAARYQHAACgiPAAAAAUER4AAIAiwgMAAFBEeAAAAIoIDwAAQBHhAQAAKCI8AAAARYQHAACgiPAAAAAUER4AAIAiwgMAAFBEeAAAAIoIDwAAQBHhAQAAKCI8AAAARYQHAACgiPAAAAAUER4AAIAiwgMAAFBEeAAAAIoIDwAAQBHhAQAAKCI8AAAARYQHAACgiPAAAAAUER4AAIAiwgMAAFBEeAAAAIoIDwAAQBHhAQAAKCI8AAAARYQHAACgiPAAAAAUER4AAIAiwgMAAFBEeAAAAIoIDwAAQBHhAQAAKCI8AAAARYQHAACgiPAAAAAUER4AAIAiwgMAAFBEeAAAAIoIDwAAQBHhAQAAKCI8AAAARYQHAACgiPAAAAAUER4AAIAiwgMAAFBEeAAAAIoIDwAAQBHhAQAAKCI8AAAARYQHAACgiPAAAAAUER4AAIAiwgMAAFBEeAAAAIoIDwAAQBHhAQAAKCI8AAAARYQHAACgiPAAAAAUER4AAIAiwgMAAFBEeAAAAIoIDwAAQBHhAQAAKCI8AAAARYQHAACgiPAAAAAUER4AAIAiwgMAAFBEeAAAAIoIDwAAQBHhAQAAKCI8AAAARYQHAACgiPAAAAAUER4AAIAiwgMAAFBEeAAAAIoIDwAAQBHhAQAAKCI8AAAARYQHAACgiPAAAAAUER4AAIAiwgMAAFBEeAAAAIoIDwAAQBHhAQAAKCI8AAAARYQHAACgiPAAAAAUER4AAIAiwgMAAFBEeAAAAIoIDwAAQBHhAQAAKCI8AAAARYQHAACgiPAAAAAUER4AAIAiwgMAAFBEeAAAAIoIDwAAQBHhAQAAKCI8AAAARYQHAACgiPAAAAAUER4AAIAiwgMAAFBEeAAAAIoIDwAAQBHhAQAAKCI8AAAARYQHAACgiPAAAAAUER4AAIAiwgMAAFBEeAAAAIoIDwAAQBHhAQAAKCI8AAAARYQHAACgiPAAAAAUER4AAIAiwgMAAFBEeAAAAIoIDwAAQBHhAQAAKCI8AAAARYQHAACgiPAAAAAUER4AAIAiwgMAAFBEeAAAAIoIDwAAQBHhAQAAKCI8AAAARYQHAACgiPAAAAAUER4AAIAiwgMAAFBEeAAAAIoIDwAAQBHhAQAAKCI8AAAARYQHAACgiPAAAAAUER4AAIAiwgMAwCz0xNPPV9fefGf18Rdfp9uhH4QHAIAh9MJrb6Y/D5dfd0v1v3+2QcN6m+5YvbLmnbQd1E14AAAYMkefet5YMHhm9Stt2+59+Imx4NC07kbbVatfWdPWDvpBeAAAGCLHnHZuWzD42e+2rla99Fpj26oXX6v+c73N27Y3/XT9eY3tnceDOgkPAABD4oLLr0uDwc9/v0318OMrGz0R2famn/52q+rJZ1anx4Y6CA8AAEPg6iV3poFgIpvvtE8jKGy49e5tP4+eidfffj89B/RKeAAAGLCl9zzUFgDC8WcsHhcMmjbado/qvY8+bez74adfVD/+9RZt28++6Mpx54A6CA8AAAP0wIqn2i78Q0yYjm3vf/xZtdX8/cZtP/OCK8b2X3jS2eO2P/fS62PboU7CAwDAgDzx9OrqR/+9WduF/4JFp7S1+eTLr6td9juyrU3449oAET0MnT9fcsd9bftDnYQHAIABeP7lNxoVklov/Hc9YFH11T/+Oa7tV9/+T3XgwpPb2mZi/YfOfaFOwgMAwAx7dc071S822K7twn/eLgdUn3/zj7R906nnXtK2T6szFl+W7gN1Eh4AAGbQux99Uq2/5c5tF/6b7bB3Y+Jz1r7TFTfc1rZvWH+rXdK2UDfhAQBghkRA2Hi7Pdov/NcGiXc+/DhtP5Hb7h5fnWnBwpPTtlAn4QFm2GtvvVeteOqZ6vrb7q4uuOL66oal9zT+rSY3wNwWQ5JiaFLrBX8MXVrzzgdp+6nEonH/vu6mbcfbY8FR1RdTDH2CXggPMEMuvmpJ9csNt2/7kO+07kbbVZddc3O6PwCz224HHtX2mf+z9edVL7y6Jm1b6vG/PVf913rtazxsv8fBU86dgG4JD9Bnb3/wSTXvD/u3fbBPZd7OB3R9JwoYL4aEPPrUM40VfE8668Jqz4OOrubvfVhq5/2OrI465dxGkH9wxZON3sLsmFDq62//Ve1/xIltn/NRnnXlqhfT9tO18vmXqp/8Zsu2419y9U1pW+iV8AB9Nt3g0BSrh0Zt7+yYwMTijuvyx1c2auDvsOch40phduNHv9qsEepPOOOCatkDj1YffPJ5em7IHHH8H9t+n3647iaN4apZ2268+Pqb1c9/v03bOZRspV+EB+ij08+/rO3DfLri7md2XKBd3NmNXoJDjjmt+s/1Nk//nur0g3U2bCzaFYtxffrlN+ljghAhtvN3577lj6dtuxG91P+98Q5t52isFfHt/6TtoVfCA/TJux992vZh3i1DJmByMRTpN5vPT/9+ZkKMN4/6+noj6BR3/zt/X268/Z60bTcaJV+32qXt+NvtfpD5DvSV8AB9ct2ty9o+0LtlAjXkrrxx6bg7roP0H+tu+l2IKKzVz9x2210PjvsdqXMo0cdffF1tsv2ebcePf8fPs/ZQF+EB+uTMxZe3fah369jTzkuPD6Mqhif9dos/pH8vw+DHv95C6B9xMSzp39bZqO33os7Vn6NnIXoYWo+/3mY7NXoisvZQJ+EB+uSsP/257YO9W4tOOic9PoyiCy6/rjFmPPtbGTb7HX5C9dnXf0+fB8Ptw8++rO5/9Mnq2pvvrM69+OrG53Csn7DZjnuPTUz+xQbbVhtts3vjIj5u8lx3y7Lq6dUvV48++XSjF6r1d6Hu+WtRLaz1+LFWxKtr3knbQt2EB+iTuoYtxUJy2fFh1Bx01Cnp38gw22L+vo1yzdnzYbhE0Ft6z0PVPoceN27htV5EidaY0J+dc7q++ef4kq/R0/X8y2+k7aEfhAfokzXvftj2Ad+tqN+dHR9GRYzhjvUXsr+P2WC9TXesXnFXeGjFEKOo0hXrLmTvX6+i6l5dAfKY084dd/xYJC5rC/0iPEAf7X3IceM+6KcjusOz48KoiDu2O+19aPr3MZust8mO1fsff5Y+Rwbj7ocem1aVrg3m7Vbtuv/CatHJ5zTKr8ZQpQULT278rGQtkXMuvip9HKXOuvDKcceMdYQ++8rQOGaW8AB9FPW3YxJb5wd+iagi424loy67YJqt4iIze47MrFfffHfcnIFOUbXozLUB4S8rV1Vr3vswPU6naBuFMmJeRHbMsPF2e1TPvvBquv9kLr/+1vR4IW4yWVCUmSQ8QJ+99f7H1Ta7Lkg/9Cey5fz9qtfffj89HoyKR574W/r3MZtd+Ocb0ufKzIjJz9n7ErZa+7kbE/JfeO3NdN/peO6l16tTz71kwqFQUWY42y+z5PZ702O02nqXA6uPPv8q3R/qJjzADLn0mpuqX264ffrB3xTbL75qSbo/jJK4k/rrGVj4LSroHHf6+Y0LvX0PO776WcHwk17VcXHK9Nz3yBMTDlHa6+Cjq1UvvZbu16sPP/2iMechO2/0bGT7tLp3+V/G7XfkCX9MC3LETac4X3YcqJPwQN+999Gn1cOPr2wsjnPy2RdXhx9/ZrXnIcdUux14VHXE8X9slDS9/d7ltd5pj8oTsYpnfGjHRLhd9juyMeEyRPnE0867tLpl2QMDGYMcz3PFU89U1992d6OS0g1L72n8ey70NMTrHtVKzrvkmsYF2YJFp7S99vElHeOE467bsF1AxcT0KLUYtdgPO/b0atcDFo097jb7HF4duPDk6oQzLqguufqmxu92lHXMjkn3TjzzT+MujuoUnz3RK9h53picff5l16b71GW73ReMOy/9E5+z2fsQ8wXi7zfbp27x+T5v5wPGPYbJeqIeXfu90Nk+5tE1ty+5475x2zfdYa/q3bXfua3HgboJDyMkLlDr9umX36TniglccWdk8532GffhNpn44LvjvofTY04lLuAuuvLG6lcbT353v1NcCNZR0Sh7fer0xNPPp+dtFXedsn17lZ2r6clnVjdqmP/nepunr+9k4v2OEJcdt5/i9zMCawSZuFuXPbbpWn/LnRu/S5dde0v1Ro9B8LW33kvfhzq9vObtsfPF/2dt6jSd1yTGmGevcV1KeveWr72ozPatS6whkJ2XemVzZjbcevfGTY6sfT/F92X0bnU+nixAxHfSj37VPuRpp70PG9cuPsda24RYeyILxlAX4WFExMTdzg+YOmQXBDE+MxbPydqX2mjbPab14R53skuqXUwmeiTioi07/lTqKss6mVgEKDt3q1icKNu3V9m5YtJf3InP2k/XTH2Zr35lTXXQ0aemj6FucXd56dov9uxxTCWqsmTHrNOJf7xw7Hzx/1mbOv3pivKx/sedvjg9Rh123OuQ9JyZiYab1CHuQmfnpD5REanzdY+//6ztTFp8+XXjHteyBx4d2x69sj/73dZt22M+xkSTorOhTXEzI76XsvbQK+FhRPQrPHT2POx1yLFpu27F8JHW43eKpfinOxl5KhF+snNNZi6Hh3U3Gn/eflXAieFs/Sg7GKu+TlVdpV8iGE23d2WUw0NcIHXeca3TQ489lZ43Ez152THqoj5//8Tww87XO8Jg1nYQOh9fTKyOKlAxvCnK+rZui8+QGP6bHacpeso6V7WO9UWsOk0/CA8joh8Xt/FB1XqOzlUv63L4cWe0nacpJrhFOdNsn17F2PfsnBOZifAQk6mzc7eK4SHZvr2IoUWt5+j3hea2uy2YcDhcN2KOTXaemRYrDZf2bI1yeMjGcddlnQ22Tc85mX6GzpgXlJ2T3sSwtM7XOoYTZm0HKeZPtT7GP+x7RCMotP4sgkTpkL8VTz1b/dd6W7TtH8N4X3zdBH3qJTyMiH70PMSHWvP4/ezeD1E/u/X5xKTG327xh7RtXaZzh3Iu9zzE0KTm8WMidNambjEuuPV5deODTz5vDB3Kjj8oP//9No05ItnjbTXK4SHWQsj2r0Pc4MjOOZmYR5Udqw4x1DI7J73pHMIaqzJn7QYtQsFkvWwxdGm6hSWeenZ19dPfbtV2nBhGHMUssvbQDeFhRPQjPMRk6Dj2gyueTLfXKYYmtT6fGM6UtatThJPPvi4bQjOXw0Pzguveh59It/dL6xjg6Xrz/Y8akwaz4w5aXCzE30z2uJtGNTx8/s0/0n3rcvZFV6bnncx9yx9Pj1WXv656IT0v3dvn0PEr+1954+1p20G76c77xz3WEAUoVq56Md1nKjFMs3PORASKZ1a/kraH6RIeRkQ/Lm533u/IxrHrmjQ7lWYd7pj0mm3vh7iIa30dJzKXhy0dfep5jWNPt3JWr6K7vfP5lYhKPTFZMDvmMHlgxcQ9W6MaHh7766p037rcdteD6XknE+PQs2PV5bJrbk7PS/eiBHdWFjUmUGftBy2GK3U+1lggMWtbKnoaOguX/PjXW3QdSKCV8DAi+tHzsGDhyY3JWNm2fogJZvFcYg5Etr0fSsdIz+Weh1jIKLrCs239duuy6V3sxVClDebtlh5r2EQPxER3nUc1PGRVaOoUfx/ZeSfT796QqPKWnZfexLypWGOm8/WOstJZ+0Hq7L3vZnhd5qU33hpXujwCREnZb5iM8DAi+hEeYrJf1KfOtvVDc2GlzooS/Xbn/Y+Mez07zeXwEHdGY0G0bFu/NXu3SsWFWHacYRXVUGL+TufzGNXwcNBRp6T71qXbVYSjEk52vDrE8LrsnPTuy79/m34mxM9iW7bPIHzzz3+1Ff/4yW+2rL78pp7HF/MqOldqjxsX0cuXtYcSwsOI6Ed4iJWhs7Gl/RLDdqIrNtvWT8efsTh9TVvN5fAQpWvrWkitG9nzzEQvRbZ/r6J6yda7HDjtxQdLLTzxrHHPZVTDQ6z4m+1bl24XzoqQlx2vLtk5qUdcmC86+Zxxr3lUOqqzqluvOstf17mIYMwB6xzK+e/rbjpjq2sz9wgPI6I/PQ+LGxe02bZ+6XdVp8wOe069qNRcDg/nXzozFZYmEnNcsufaKia2x3oU2f7diqEDneeOYVHRE5O170XnMIJRDQ9xxzXbty4fff5Vet6p9Hvy/XRW36Y7WaW4uCkS8yOy9jOtc27NwpPOTtt1K9ZE2ni7PdrO8cNfbmKlc7oiPIyIfoSHURGTzrLXtNVcDg+DVlJ1Kavr3ovdD1yUnqfp6iV3pvt1a/s9Dm47/qiGh2y/OsX8hey8U4m1TrLj1eVvq19Oz0u9Yv2eztc+1lWI78es/UzbZPs9xx5X9HZlbXoRNz+yXuTbu1wJn9ElPIyIWLWy8wODMj9YZ8P0NW01E+GsZPJ2VOjI9p3Nrrt18gX7YlhC55jeXpWUNKx7GFPrePwZCQ9n/mnsfMMQHmLuR7Zfnb76xz/Tc0+l38P2Hn3qmfS81C9uRsQd99bXP/6Wp7ueQj/EEMbWxxWrrWftehG9b53DA+M7ToBgOoSHETGTPQ9R4eKuB1c0yoaG8y+7Nm1XhxhOEHWym+eKC81+TKjOJrW2GsWehxhDe+WNS8de+3uWP1ZttG17t3gdLr3mpvS5NsWQn2y/bv1uq13S83SqexJ5DMlrHnsUex7e/ejTdL86ZectEevMZMery32PPJGel/6Iz8lYR6H1PYiF5Qa95kZnz0i/Hk/M9dhxr0PazvV/fr5hdcuyB9L20El4GCFxgRcfmnHBF+VOJ1vZsltRLSU7dz8CxBbz900XcYs7S1n7Xnz46RfjztMpXtsblt5THX78mekxelWyzkM8znifowcixubHuP3sWL2KKkjZZMP31l4AtlYNqUMMSeo8T6u6f7cOOvrU9Dydbp5gcadutS6E+Npb7429j9ffeld1yDGnpfv0ojU8vLzm7cb5YgLlNTfdUe19SP2FEKYKD/G7k+1Xp+y8JWJYWXa8uty7/C/peemflc+/NG4htbjxFIthZu1nQizu1vp4ltxxX9quLtlq7lctuSNtC62EhxG29J6Hxn1w9OrZF15NzxVdpVn7XkSPQ3au8Juah7GUhIdWMdG27gotJeEh04/a+ZONEb6g5vNNFR6yL8BeRBjJztMpxqln+/disjH5dRcLaA0PmbpXcZ8qPPR7PYWQnbdEv3se4sZDdl7668XX30w/pyf7bumnzuHF3ayIPl17HnJM2znDZdfekraFJuFhhMXkqc4PjV7EuPPsPE0zeUGffSD2YrrhIcRqttmxulUybCkTXd/Z8bq17W7f3yHPLH98Zbpfty668sb0PE3rbbZTul+3YuhAdp5O/RgKGKWIs3OFGBef7dOtqcLDbXfXe3NhGCZMZ+cs0e85Dyufs+rvoET53s5qWjGE5883LE3b91NUfmp9HK1DGfsl5gFlJddLb6IwmoSHEdf5gdGL5iJuE4lhRtl+3YhFm7JzNB1a813TbsJD3eGs256HGAKTHa9bUTM9O09T3Xfkp+p5yPbpxXSGkGT79yLmjWTnCXGXNNunW1OFh7rDSkl4+Pnvt0n3rUt2zhKtVXD6Icp0ZudlZsTne6zl0vm+nHbepWn7fvn6238N5Pxx3mxo5ElnTf4ZwegSHkZc54dFL+Juf3aOpq1qvHs3VS9H3fMOugkPITtWt7rteag7PJx09kXpeZqiUlG2X7cm63lY8179E9VjnkF2rky2fy9ivkF2nlB3T8dU4eHJZ1an+3WrJDz0e5G4bv+O6+417ZSdk5kV8+f2POjoce9NzA/M2vdL67lPOefitE2/LDpp/GJ6AgQZ4WHEdX5Q9CK6PrNzNNV5YRBf5tk5muZieBiWnoepvtDqDg+T9TzERN9sn150Ltg2mVh9OjtGty7888QX2HVX9BrG8HDAkSel+9alda5OVHdaeu/yxtCQBQtPrnba+7Bq/lrRaxnzwT787MuxtlGJJzteHWLITPM8DF4212evg49Oi3PULX7nWs870+EhRFhofQzh2NPOS9syuoSHEdf5IdEL4WG87FjdEh7GizkC2T69mE55xFh7IztGt866cOIJkqMQHs684Ip037qsevG1RpWw6D3LtneKyfjX3lzvgoCdIjBlrwWDEevGHH/G4nHvU1Tcmqpkd69irYnWc04136tf/pj8HcawpnhtsvaMHuFhxMXEsM4PiW7tdcix6Tma6pzzMNPDlj7oMjz82zobpcfrRskicZkYT50dr1tTdWN3lhvs1WR34/vR8/DYX1el58r85DdbpsfoVlSqys4T6l7osXWRuMwTT898eIg5H9m+dYmLsVjHI9s2KFfccFv6WjBY8bnT+V5ttsPejQnWWfs6dM4zuvP+R9J2MyEqLrU+lnDgwpMb8yOy9owW4WHEdX449GKq8FBnxZJR7Hn4RZfhYbbPeZgsPMQXebZPL2Ktg+xcmWz/Xkx2IflG3eFhip6HuhffKwkPndVmRsFE5a0ZvFhnofP9+u0Wf6heeuOttH2vbrz9nrZzlax030+XXzc+QOx3+AlpW0aL8DDiOj8YejHVhOk6hy1Fec7sHE219zx88nl6nsnEGNnsWN362frz0vNMZaZ7HmYyPIRsn15MZ5GobP9eTHansfZhS1P0PDyw4ql0v26VhIew2Y57p/vXLXpCr7zx9sbieHc/9Fh1wRXXV5vusFfatl+6vSHAzLlv+ePVv6+7adv7FlXBooc1a9+LhSee1XaeWB8pazeTrr/t7nEjFPZYcFT15d+/TdszGoSHEVb3xchU9f/rDA+xOnZ2jqa6w0OUyczOM5m6S12GbFXnqax46tn0WN2Ksa/ZeZrqDg+nnntJep6muivhxN2/7Dyd+rEicozJz84V6v57jb+R7DxNcbGf7detyeZztKr7vJmYpDzRBNgYOtW58nC/mIg6Ozz+t+fGFUf48a+3qH1xv9b1JmLxuqzNIGSr6cd8oKwto0F4GGGxCE7nB0IvoiLJJ19OPKGs7jKMky2oVXd4uGXZA+l5JrP/ESemx+rFsgceTc81mdLJoaWmqg5Td3iIKjjZeZp23u/IdL9uTdXT0dQ5ubEOk1V0qTs8xNoF2Xma1t9y53S/bu172PHpeTr1Y/G9TlFNKTt3Uwz1m4m5EdMpC8xgrXrptUa57Nb374e/3KSrz+RM57pAU4X7mRa9c//2i43bHuOOex0yI1WoGD7Cw4iKhbw6PwjrsPuBixp3urM5AnWHh5i8Fgt6ZRPY6g4Pv9p4+8b417io6DxXJptsV4dYHC8WaYtx+W9/8El67lYx9CR6abJj9SJqn0dVoiws1h0eQozPj+NmX1TnXHxVuk+3Fp509rhzZGJ4U7Z/t7bZdfKeu7rDQ5hoQbx+/f6eufjythKoE6l7hfhOTz27Oj1vqwiHP/3tVun+dZhq3hbDJz7/Y85D6/sYQ3pKV6WfTAwPaj3urcseTNsN0v2PPtkITK2PM0YcdNMjzuwmPIyIflx4lIjzNh9D3eGhU1woN89Vd3jIxMVs83yDen1DDJ9pPo4Yv5216bfdDjxq7DH0Izx0ar0rV/ewrHm7HDB27Mmcf9m16f7divKI2Xma+rEgXliw6JRGWIjfnbiLGoUPsnb9cPSp+bCdlc+9mLavS8xvyM7bKu40111Nq9VVSyZeEJDhFeuDZJUDz7/0mrR9qT/se0Tb8aJ4QNZu0JY/vrLtcYap5jsy9wgPI2ImhgJkWsNDnStMZ+KuSPNcoxoeYgxu1qbfBhkeQt1DbErmuLSOT67D6lfWpOdpismT2X6z2UThIfQ7xFx549L0vCEqTfVz3sO6G3W3WjzDIdZ7iOGUne9rrA+RtZ9K542B+K7M2g2LrArTsIYd+kN4GBF6Huqn5+F7gw4PsZBc1q5brc8nc/WSehcOi7HD2Xk6ZfvOZpOFh1jXIobpZfvV5T/W3bQx8fOY085tTMyPKjKdw1L6oXWcfAz5iL/bpsnmjTFcomxp53t70NGnpm0n0zlMMKqAZe2GQRR16BzOF383WVvmLuFhROh5qJ+eh+8NOjzEBVeUss3admuilbT78Rrf90hZediY5Jzt329x3g3m7ZZu68Vk4SFE70C232wWQaX1Of5l5aq27VP1QDE8YsG07Ltm/j6HF88DiJKnUVmpuW9Ucfrsq+GchBw9slGmtvW5xmdDt+sgMXsJDyNCz0P99Dx8b9DhIcTqzFnbXsTv7NkXXdmYlB1ifkDWrhc77FnW6xAOO/b09Bj9FnME9jr46HRbL6YKD6Ef5x2UeTsfMG7Sf/RCtLaJHpfW7Qy/My+4ou09DLEoaslQns6VnE8779K03aDFDcj/3niHtse6/la7VO9+NHXhDuYe4WFE6Hmon56H7w1DeAhRgStrP8yiqk/2XDJR1SU7Rr899tdVAwsPH3z6Re1reQxCDIeKybadz+/Sa25qa9f698zsEaXPOxdT+/1Wuza+e7P2IXonWufW/GCdDdu+M4dFBIQICq3PLYLEZM+NuU14GBHCQ/2Eh+8NS3h47qXX+1Katl+uv/Wu9HlMJC6ks+P0U5QpjnMPKjyEeF9/ueH26TFmg3U22HbCkLhg4cltbYd1yApTi/Kqre9lWG+THSd87zvLTE+16vsgxJCkzpsyMXTphVcNrxtlwsOIGNTFbWt4MGypPwxbajeo12C6YphV9vinkk3S7KfF//9xDjI8hLgBstG2e6THGWZRlWui9WFirk5nOdisHbPHQ4+NX1snFlCNdXFa2618/qW2RdeiAtewTZaPIBulq1ufS8zJePaFV9P2jA7hYUToeaifnofvDVN4CFGLvN+VenoR8yiyx10ixsTHF3h23LrFJPTmuO1Bh4cQC8xFZZfsWMNqsoUlY/J0a9uNt9sjbcfsEkEhAkPrexuBIoJFbI+/qeiRaN1++73Lxx1nkL785ttx5WijOlmUMc7aM1qEhxGh56F+eh6+N2zhIaxc9eK4L+hhcMUNt6WPdzpuu/uh9Nh1i/M0zzkM4aHprD/9OT3eMIqJsxHqWx9//M2efv5l49ruc+hxbe2YvWKoUgz563yPb7vrwXEX5a2fn8Pgq2//Z9wq79FL0gw/IDyMCOGhfsLD94YxPIQYr3vQUfVXSOrGpjvsVT29+uX0cXYjKiBl56nLwhPPajvfMIWH8PKat/tS/apfIkScdPZF1c77HZluDxOVB2Z2ih7/mDSdvddNETCyifSD8s0//9VYq6L1McZE7jvvfyRtz2gSHkaEYUv1G8bwYNhSLsYX73rAovRY/bbh1rs37jZmj6tX2Uqvddj3sOPHnWvYwkPTqpdeq/Y/4sT0+LPNsA1doXcxRCmCY/Z+//u6m9Z6Q6EOnUPpwo2335O2ZXQJDyNCz0P99Dx8b9jDQ1N8Uc/U3ertdl9Q3XHfw+njqFNUPdnrkGPTxzBdMS77kqtvSs8zrOGhKRawiqFArQtuzTYfff5V+tyY3aIkaywc1/l+L73n+2GBwyAbDhg3KLK2jDbhYUTEl1JzoauZ1PpleO3Nd6Zt6hLDGJrnuuvBFWmbOr39wfeL4wzq9Q2tpR1jcmbWpt9uvvP+sccQr0vWpk7x/jbP1434Mo+7vMeedl5toTbq+EclpFj0KXr6svP2U1RAOXPx5dVmO05/rYvt9zi4uvDPN1Rvvv9ReuwQZSiz96IX9y1/PD1Xrx5dG6LjQigW4Mueb7c232mf6uBjTh1XgaYOMQ4+ey7MHa03Lk4995K0zaDETYPW38cQf0NZWxAeANZa+dyLjUXYzlh8WXXosadXu+6/sHEhHiUU44s0Kg9F2c35ay/yDlx4cnXCGRdUF1+1pHrkib811l/IjjkoMYY6hvFdeePSsecTY+3jsTfF44/VjWNeSHaMuSSCVdzlPffiqxsXcBMFxlj4Kmrax2sVIeHksy9e+xre3ph833q81a+sSffvxU0tAZy566old8xIj+R0XHfr+MUnjz9jcdoWgvAAANOUVUvqVqygnZ0D+i0mQseE6Nbfx5gwHROns/YQhAcA6MK8nesZvtQ67A9mSpRebV2oLkSJ1ijVmrWHJuEBALoQlXSimlbrxdd0dTv5H3oRi73Fom+tv4sx7yYWh8vaQyvhAQC69O5Hn1SbbL9n20VYqVgtOzvmsIpCA9nE99nm6iV3NuY4Zc9xFMQcoM5V6qMIQGvxDZiM8AAAPfjgk88bE9BbL8amEtWtsmMNs5iInz2X2Wy73Q+qVjz1bPp856Ioafzz32/T9hpE+B2FwgnUR3gAgBpEid6oUrPRtnu0XZyFqNYVa2VERa/WMs+zyTsffjzuec0F/+fnGzbWHPn4i6/T5z1XxO9nVBRrfe7rb7VLo/csaw8TER4AoA/WvPdh9eQzq6u33v843c5gRGnlhx9f2ej9aZZiDutttlOjFynbZ7aLgBBBoTU4RJAYxJo0zH7CAwAwpeh5OPKEP84ZZ190VaO34bjTz2/0PsQF9b6HHZ8+99kshiR1zsuJoUsxhClrD1MRHgCAKc21YUsRHprP7cwLrhj7+VwqnRuToDtXRI/J0jFpOmsPJYQHAGBKER62mL/vnNEaHr76xz/H5qrEKuOtz3u2irKrUX61NThEedYo05q1h1LCAwAw8m68/Z6xi+zZPnk6FnqLBd9ag0MsCBcLw2XtYTqEBwBg5L321ntjF9r3LH8sbTMbfPPPf1UHHX1qW3D4wTobVnfe/0jaHqZLeAAAphTDlmJi8VzROmypqVnK9JRzLh63bbY45rRz24JDPNfoVcnaQjeEBwBgSnN5wnTTRtvs3th23OmLx22bDU7844XjnmesqJ21hW4JDwDAlCI8XHbNzXNK6/P77Ou/j11w37B09t2p/9MVN7SFhnDuxVenbaEXwgMAMPKeeHr12EX3qhdfS9sMq+hdaA0NIXohsrbQK+EBAJhS9Dxcf9vdc0rr89vn0OMaF90//OUm1dff/qtt2yBdf+td1a4HLKrm731Yaqe9Dx0XHEJnu1PPvaTxHmbngOkQHgCAKc3lOQ+3Lntw7OennXdp2/MepJXPv9T2mHs1TM+N2Ut4AACmNFfDw213P1T95DdbNn628XZ7jHveg7Tk9nvHPe5e7Lr/wvQ8MB3CAwAwpVg4bS559Klnqj0POrrt4vq5l15Pn/ugCA8MI+EBqN0Hn3zekG0DZqcPP/1i3Dj62WizHfced1H9m83nV088/Xz6vAepMzw8+8KrabvJbLL9nmP7Cw/UQXhgoJY98OhYXe1fbLBtdf6l16TtunHfI09Uu+x35NiHZohVN59e/XLavtQlV99U7X7gonRbmLfzAdWSO+5Lt03HZ1/9vbrs2luq9bfceezxr7vRdtV5l1xTvf/xZ+k+0/H4356r1ln7mmfberXHgqMasm3disoh/7ne5o3XYd4uBzTuGmbtpvLR51+NHSez3+EnpPtN5fDjz6xOP/+ydNt0xTCKbXZdMPaY/mPdTRsLP7285u20/SC8+ua71X+tt0X17kefptunI1b2Pf6MxY3jNZ/zVvP3q25Z9kDavtSehxxTHXH8H9Nt5192beOzJ9tGbq4NWwqxgNpxp5/fKNOaPedBEx4YRsIDA/PQX/763YfZAYuqR574W6MSRPz7rAuvTNuX+vybf4xVzdhxr0Oqe5f/pVqx9kLztrserDbc+rugctLZF6X7loiL2MkuOn654faNC5NsW6nVr6yp1ttsp8ZjjYuqePwhanb/6FebVT/73dbVX1auSvctFReocfxPvvw63d6tKHEYxw11lTs8/LgzGse74IrrqwdXPFltvcuBjddh1UvdHb/5eoZYUXbBwpPH/v38y2+k+0wl7mguWHRKuq1UhMLtdj+o8Vz3Pez4xt9IPKbrbllW/Wrj7Rs/v+KG29J9Z9qTz3xX1rLXYR4RtOM4v9hgu+rKG29vPN/lj69svCfx8whR73UZUCLoxzGeWf1K28/jIvhH/71Zdexp57X9nMnNlZ6H+N268M83NP6+hr2HVHhgGAkPDMz8fQ5v3F1s/dnJZ1/c+IBr/dl0LTzxrMaF5UR3/xdffl31099uVX365Tfp9qn0OzzEnfHobYgP/OxCds07HzRC0U/Xn9f4/87tpfoVHhaddE613qY7NoYBHHlCftd3Ota8+2HjcV5+3S1jP4uAGBf9cbe/tW03frfVLtUp51ycbpuOxkVJj+HhD/se0eiBi4DUuS16ok4880/VBvN2G7dtEOoIDyueerZxjIOPObX68LMvx22PEBGhIj4rOreV+u0Wf6h2WvvetP5s0cnnNHo5lK1k2N105/2Nv5Gmzhsm7370SfXCq2saPckRkrPvhLbwcMDEveZQSnhgYNZfe9EWYaH1ZzEUJT7guh0KER+gsf/Naz9ws+116Hd4iAvZCDdr3vsw3R5isl9cnHc7xCb0IzzEnfM4ZgS05mqnvd7Zi3HIcZzOu8cRTHY7sPehUcMSHpp3GON3ONs+bOoID/F3FEMLs21Nzb/peH2y7VNp/p7fs/yxxr/j4iv+fcHa39HOtjBsOleN3nynfarfb7Vr9bP157X9vFP0rP167XfElvP3a/x/8+fRazusQ7SYPYQHBibGI8cd9hdff7Pt53/rYU5CXFD2u9Rev8PDz3+/TdHQrVjgKL4Msju2JfoRHuKCLI4Zd8NC/H9rj0E3msfpHLseQ1l66XlpGpbwEHfH9z7kuHTbMOo1PDRDYdSxz7a3itell96HmCMTQxbj/6N3J3qtoveqsx0M2pd//7a6/9Enq6NPPa/xexp/I3WLOVR7HXJs4ztEYQu6ITwwMNHVGmP348MsxvXH8JSs3XRstsPejWNl2+oS4aHzw7hTt+EhJo7G/g8/vjLd3uqVNe802sbQj2z7VPoRHmJc/iHHnDb27xiOUkeYa45djyEo1926LG3TrWEJDzGJ+9Jrbkq3DaNew8Pl19/aGDqUbesU73/0xmXbSjSHR+1/xImN/9ZR0GAUxIXlAyueavQkNgLc2t/xURavRfY61SGKIcR3y1Q9CnX7t3U2ary3EVi++efwrKrNcBMeGKgY3x+TYKN7NT7I4kIz7jRnbUvEUKi6Kt5MJD7go3fgnIuvSsVFYLfhIeY4xOvwxNOr0+2t3v7guzvyUVUq2z6VusPD0nuXj3vsMak7fvbok0+3te1GvDaHHXt643ihl0nvrYYlPMRzuuamO9Jtw6jX8BDDMUqrfcXrEufKtpWKCehxjC3m75tu53vRmxnzQqISUbxmfOfG2+9JX69exPymmMOWna9T9NTvdfDRjVWiY8J3FBiIORF3Pbii8Rl778NPND7Xo8BC3IiI76OoMBhDnbLjdYqhsBdftSR9nNBKeGBoxMVhfHhtu9uCdHuJGD9dd4nQTv0cthTVTOJD/Ppb70q3t3rosacabbutOFR3eIiqONHz0Kxa1BSvRy9zMzJx1zoe+9VL7ky3T8ewhIe4MOh3r1mdeg0PMS8p9i+ZtHzCGRc0ep2ybaVieFSc7/a1ITfbznfufuixRknoeK1aRXGL1rvwo6jOnofoaW8G2onE90wMBY2hvHUMs4vejeh1i0nT2fmaYohfFCvIjgFBeGBgojRr5wdUlFOND6833/+o7eelmpUpVq56Md1eh37PeYhSnc3x2ZOJMfLrbbJjuq1EneEhJjM3v3gmEusCZPtOJS4uY35M5yS/CIlTTbYtMSzh4eyLrmxUCatj+N5M6DU8NHvOooJUtr0pPguiN6/XsqrNIYEuiiZ25/2PjP29hqic1k1pUCYXvQatr3OrqC4WveedBSLqFj380TsRQ32zxxEOOPKk6oNPv0j3Z7QJDwxMlI/rLOXZrDjTy8JTMSwhFmqb6E5NfBjG4mvZthJxsTNZeIgP/14Wu4v5DvEaxJoO2fbQDEkx4S3bXqLO8BBDHGJyX7YtxLZuhxnFXbd4nNGL0frzmEBbRy9TDHWrLTwsPDndViJKsUbPTdyNzLaHuADux9CJ6WiWD+41PIT4W4pjrHxu4rAfFzARqt54+/10eynhYXIR0pvj7eOmRPPvLQLxdrsvYK1eq/jF91rnwqVNUUEpelK//ObbdN9+ivUuopBA9rjid+Gvq15I92N0CQ8MTHMSbHPCb5QmjWFLscBbZ9vpaA5PiLv3sUBc8+exrkOEk7i4jy/JYV3nIcSdp3gOB669GG29OHt97QVUDG2Jbb1cqIa6wkOMj46Lu5hUmW0P8Xr85Ddbdn2uGDIRJQdjWFf8u9lDFSuUd7adrmHpeQgxdjmeVwzda11BO573ZdfcXP3411tUG23b32piU4m/nStvXNpY2DEeaxQ+yNqViN6keF/jODE8o3Xl9KjGtMOe340Fj7Dcul83hIfJRdnjeH1+sM6GbRWw4nc6fk5vJcBjLlh8N3QeM+b73brswXSfmfb06pcbPdqdjzFctWT2zMei/4QHBipqTrd+QMVE5PiSz9pOR9wpifKMrcduisoSvZzjpLPWhodJLuBivHAdNeTjAi2CTufjj4vw8y7pvmejqTnBudN051BE9aPYb7KJ7nFXM9qUzOXIxJ3p1lrlYbI79NMRd/xiAmK2bTqiJyQmJ2bbpiOqnkQ4bX2uTbHS9qBLKx501PcXk5vusFfaZjpizZLouWp9nk2xIF5MAs32m67ouYhjRujJto+yl954a+w1P2Nxe8GJeL/5TrfhIYZ+dX5+hVjQtNubWP0SFZeivHbcEOp8vPGdlO3D6BEeGLil9zzUqAoRlSOad5brEusAxN3pGGMak93qGL8ZPQHNBacycUe/2/H9mRi2ExfoV9xwW2OoSNamG/Fat1aJChFKpvtl9uqadxrvYbatVfT69PK6xNj36K2Kx9lthalM/H688Fr7WiPdiOFmdXbvx+TGCHhR/SQqqdQ1sb1XUSHthqX3NH4fu52blIlhW/E8Yxz2LcseqFa/0n2PxkSi98b6DuPFne/mBWK8v1kbuhOfLc2S5E1RZazOz7B+iMcdQ4BbH3cYll4SBkt4AIARFkMx48IwKn51bosbL3xnuj0P0eMa85haL76jd70fwbgf4oZFVNFrffyhjtLbzG7CAwCMsJivExeFMTm9c5s5D9+bbnjoLIkaPQ519HLOpCxARCDSQzXahAcAGGERGuKiMEJE57YIDz9cdxPWmk54iIXaWi+4YzX1XooLDFIEiM6F5upeu4fZRXgAgBEWBR7igjDW08i2Mz1RObBzgnSsoZG1nS1efP3NtucT6qh2x+wkPADACIsKX80Lwjqq3Y26407/rpx208KTzk7bzTZRLa/1eUU1tKjOlLVlbhMeAGCEffHNPxrrDcQF4dGntq/kHcOWYkFP9iwatvTW+9+VpW6K13UuVfja6+Cj257fHfc9nLZjbhMeAGDExRCU5gVhlAhu/tyE6e+VhIeTz764bZ8/3zC31kaIFeZbn99mO+6dtmNuEx4AgGqnvQ9tXBDG5OBYT+XLv39bHXXKufx/JeEhKio1L6xjNfZYRT1rN5vtf8SJY88xTHdhUWY/4QEAaNxVjqpAzYvCWOk8Vht+7K+rGiuBZ/vwvWdWv9J2UR0X2SueembOiQUiW5/n4suvS18P5i7hAQBoiFX5t9/j4LaLQ75z4+33pK9ZU1xEZ/vNdVvN3y99PZi7hAcAYExU0Lns2luqH/2qvdzoqJsqPBx67OnpfqMgez2Yu4QHACD13EuvV3c9uKK6+KoljXkQoyyGJWWvUdMlV9+UXljPdVvM3zd9PZi7hAcAgBo8+czqNHjMVdfduix9HZjbhAcAAKCI8AAAABQRHgAAgCLCAwAAUER4AAAAiggPAABAEeEBAAAoIjwAAABFhAcAAKCI8AAAABQRHgAAgCLCAwAAUER4AAAAiggPAABAEeEBAAAoIjwAAABFhAcAAKCI8AAAABQRHgAAgCLCAwAAUER4AAAAiggPAABAEeEBAAAoIjwAAABFhAcAAKCI8AAAABQRHgAAgCLCAwAAUER4AAAAiggPAABAEeEBAAAoIjwAAABFhAcAAKCI8AAAABQRHgAAgCLCAwAAUER4AAAAiggPAABAEeEBAAAoIjwAAABFhAcAAKCI8AAAABQRHgAAgCLCAwAAUER4AAAAiggPAABAEeEBAAAoIjwAAABFhAcAAKCI8AAAABQRHgAAgCLCAwAAUER4AAAAiggPAABAEeEBAAAoIjwAAABFhAcAAKCI8AAAABQRHgAAgCLCAwAAUER4AAAAiggPAABAEeEBAAAoIjwAAABFhAcAAKCI8AAAABQRHgAAgCLCAwAAUER4AAAAiggPAABAEeEBAAAoIjwAAABFhAcAAKCI8AAAABQRHgAAgCLCAwAAUER4AAAAiggPAABAEeEBAAAoIjwAAABFhAcAAKCI8AAAABQRHgAAgCLCAwAAUER4AAAAiggPAABAEeEBAAAoIjwAAABFhAcAAKCI8AAAABQRHgAAgCLCAwAAUER4AAAAiggPAABAEeEBAAAoIjwAAABFhAcAAKCI8AAAABQRHgAAgCLCAwAAUER4AAAAiggPAABAEeEBAAAoIjwAAABFhAcAAKCI8AAAABQRHgAAgCLCAwAAUER4AAAAiggPAABAEeEBAAAoIjwAAABFhAcAAKCI8AAAABQRHgAAgCLCAwAAUER4AAAAiggPAABAEeEBAAAoIjwAAABFhAcAAKCI8AAAABQRHgAAgCLCAwAAUER4AAAAiggPAABAEeEBAAAoIjwAAABFhAcAAKCI8AAAABQRHgAAgCLCAwAAUER4AAAAiggPAABAEeEBAAAoIjwAAABFhAcAAKCI8AAAABQRHgAAgCLCAwAAUER4AAAAiggPAABAEeEBAAAoIjwAAABFhAcAAKCI8AAAABQRHgAAgCLCAwAAUER4AAAAiggPAABAEeEBAAAoIjwAAABFhAcAAKCI8AAAABQRHgAAgCLCAwAAUER4AAAAiggPAABAEeEBAAAoIjwAAAAF/m/1/wAap16C525sXAAAAABJRU5ErkJggg==";

const FONT_IMPORT =
  "@import url('https://fonts.googleapis.com/css2?family=Noto+Serif+SC:wght@400;600;700&family=Noto+Serif:wght@600;700&family=Noto+Sans:wght@400;500;600;700;800&family=Cormorant+Garamond:ital,wght@0,500;0,600;0,700;1,600&family=Inter:wght@400;500;600;700;800&display=swap');";

const COLORS = {
  paper: "#FFFFFF",
  paperDark: "#F2F2F0",
  card: "#FFFFFF",
  ink: "#16181A",
  inkSoft: "#33383B",
  metadata: "#8A8F93",
  seal: "#31708E",
  sealDark: "#24586F",
  bamboo: "#2C5A78",
  bambooDark: "#1F4258",
  gold: "#2C5A78",
  grid: "#C7C9C5",
  hairline: "#E3E4E1",
  chipBg: "#F2F2F0",
  error: "#A8482F",
};

/* ============================================================
   UI TEXT TRANSLATION — covers the highest-traffic screens
   (tab bar, header, Ghép bộ thủ, Flashcard, Luyện viết). This is a
   first pass, not full coverage: deeper/rarer text (admin panel
   internals, some less common messages) is still Vietnamese-only
   for now and can be added incrementally the same way. Values can be
   a plain string or a function for sentences with dynamic parts
   (e.g. counts). Falls back to the Vietnamese string if an "en" key
   is missing, so nothing ever renders blank.
   ============================================================ */
const UI_TEXT = {
  // Tab bar
  tab_home: { vi: "Trang chủ", en: "Home" },
  home_headline: { vi: "Bạn muốn học chữ Hán bằng cách nào?", en: "How do you want to study Hanzi today?" },
  home_subtext: {
    vi: "Luyện tập bộ thủ, ôn flashcard, và luyện viết tay —\ntất cả tại đây.",
    en: "Practice radicals, review flashcards, and master handwriting — all in one place.",
  },
  home_signup_button: { vi: "Đăng ký miễn phí", en: "Sign up for free" },
  home_status_lookups: (used, limit) => ({ vi: `${used}/${limit} lượt tra cứu đã dùng kỳ này`, en: `${used}/${limit} lookups used this period` }),
  home_card_radicals_desc: { vi: "Ghép chữ từ bộ thủ →", en: "Build characters from parts →" },
  home_card_flashcards_desc: { vi: "Ôn tập theo phương pháp lặp lại →", en: "Spaced-repetition review →" },
  home_card_writing_desc: { vi: "Luyện thứ tự nét bút →", en: "Practice stroke order →" },
  home_hw_step_trace: { vi: "1. Nối điểm", en: "1. Trace" },
  home_hw_step_recall: { vi: "2. Nhớ lại", en: "2. Recall" },
  home_hw_step_reveal: { vi: "3. Đáp án", en: "3. Reveal" },
  tab_play: { vi: "🀄 Ghép bộ thủ", en: "🀄 Combine Radicals" },
  tab_flashcards: { vi: "🧠 Flashcard", en: "🧠 Flashcards" },
  tab_writing: { vi: "✍️ Luyện viết", en: "✍️ Handwriting" },
  tab_add: { vi: "Tạo thẻ mới", en: "Create New Cards" },
  loading: { vi: "Đang tải…", en: "Loading…" },
  tab_radicals: { vi: "Bộ thủ", en: "Radicals" },
  tab_hanzi: { vi: "Hán tự", en: "Characters" },
  tab_vocab: { vi: "Từ vựng", en: "Words" },
  tab_library: { vi: "Thư viện", en: "Library" },
  tab_management: { vi: "Quản lý", en: "Management" },
  mgmt_sign_in_required: { vi: "Vui lòng đăng nhập để quản lý tài khoản và thư viện của bạn.", en: "Please sign in to manage your account and library." },
  mgmt_sign_in_button: { vi: "Đăng nhập", en: "Sign In" },
  mgmt_account_tab: { vi: "Quản lý tài khoản", en: "Account Management" },
  mgmt_library_tab: { vi: "Quản lý thư viện", en: "Library Management" },
  mgmt_messages_tab: { vi: "Tin nhắn", en: "Messages" },
  msg_empty_state: { vi: "Chưa có tin nhắn nào. Gửi tin nhắn đầu tiên cho quản trị viên bên dưới.", en: "No messages yet. Send your first message to the admin below." },
  msg_placeholder: { vi: "Nhập tin nhắn…", en: "Type a message…" },
  msg_send: { vi: "Gửi", en: "Send" },
  msg_from_admin: { vi: "Quản trị viên", en: "Admin" },
  msg_from_you: { vi: "Bạn", en: "You" },
  admin_nav_messages: { vi: "Tin nhắn", en: "Messages" },
  admin_messages_title: { vi: "Tin nhắn từ người dùng", en: "User Messages" },
  admin_messages_pick_user: { vi: "Chọn một cuộc trò chuyện để xem", en: "Select a conversation to view" },
  admin_messages_none: { vi: "Chưa có cuộc trò chuyện nào.", en: "No conversations yet." },
  admin_messages_search_placeholder: { vi: "Tìm người dùng theo email…", en: "Search for a user by email…" },
  admin_messages_new_thread: { vi: "+ Bắt đầu cuộc trò chuyện mới", en: "+ Start new conversation" },
  mgmt_tier_label: { vi: "Gói hiện tại", en: "Current Tier" },
  mgmt_lookup_usage: (used, limit) => ({ vi: `${used} / ${limit} lượt tra cứu đã dùng`, en: `${used} / ${limit} lookups used` }),
  mgmt_course_label: { vi: "Khóa học:", en: "Course:" },
  mgmt_upgrade_title: { vi: "Nâng cấp gói", en: "Upgrade Your Tier" },
  mgmt_upgrade_body: {
    vi: "Xem bảng giá chi tiết trong mục Bảng giá. Nhấn nút bên dưới để yêu cầu nâng cấp.",
    en: "See full tier details on the Pricing page. Click below to request an upgrade.",
  },
  mgmt_upgrade_button: { vi: "Nâng cấp lên Premium", en: "Upgrade to Premium" },
  mgmt_my_decks_title: { vi: "Bộ sưu tập của tôi", en: "My Decks" },
  mgmt_my_lists_title: { vi: "Danh sách của tôi", en: "My Lists" },
  mgmt_lists_description: {
    vi: "Đây là các danh sách bạn đã đặt tên cho chữ, từ, hoặc bộ thủ do bạn tự thêm vào. Danh sách chính thức (dùng chung) chỉ quản trị viên mới có thể sửa.",
    en: "These are the list names you've used on characters, words, or radicals you've personally added. Shared/official lists can only be edited by an admin.",
  },
  mgmt_no_personal_lists: { vi: "Bạn chưa có danh sách cá nhân nào.", en: "You don't have any personal lists yet." },
  mgmt_make_public: { vi: "Đặt công khai", en: "Make Public" },
  mgmt_confirm_make_public: {
    vi: "Đặt bộ sưu tập này thành công khai? Mọi người dùng sẽ có thể xem và sử dụng.",
    en: "Make this deck public? Every user will be able to see and use it.",
  },
  mgmt_rename_list: { vi: "Đổi tên", en: "Rename" },
  mgmt_delete_list: { vi: "Xóa", en: "Delete" },
  mgmt_rename_prompt: { vi: "Tên mới cho danh sách này:", en: "New name for this list:" },
  mgmt_confirm_delete_list: (n) => ({
    vi: `Xóa danh sách "${n}"? Các mục bên trong sẽ không bị xóa, chỉ gỡ khỏi danh sách này.`,
    en: `Delete list "${n}"? The items inside won't be deleted, just removed from this list.`,
  }),
  mgmt_item_count: (n) => ({ vi: `${n} mục`, en: `${n} item${n === 1 ? "" : "s"}` }),
  suggest_revision_button: { vi: "Đề xuất chỉnh sửa cho quản trị viên", en: "Suggest a revision to admin" },
  suggest_revision_placeholder: { vi: "Thông tin này sai hoặc thiếu điều gì?", en: "What's wrong or missing here?" },
  suggest_revision_submit: { vi: "Gửi đề xuất", en: "Submit" },
  suggest_revision_thanks: { vi: "Cảm ơn bạn! Chúng tôi sẽ xem xét.", en: "Thanks! We'll take a look." },
  tab_premium: { vi: "Bảng giá", en: "Pricing" },
  tab_blog: { vi: "Blog", en: "Blog" },
  tab_about: { vi: "Về chúng tôi", en: "About Us" },
  tab_feedback: { vi: "Góp ý", en: "Feedback" },
  tab_privacy: { vi: "Privacy", en: "Privacy" },
  tab_terms: { vi: "Terms", en: "Terms" },

  pricing_message: {
    vi: "Chúng tôi luôn cố gắng mang đến những công cụ học tập tốt nhất cho cộng đồng với chi phí thấp nhất. Chúng tôi đang xây dựng bảng giá. Trong thời gian này, vui lòng liên hệ với chúng tôi nếu bạn muốn nâng cấp tài khoản — chúng tôi rất vui lòng nâng cấp miễn phí cho bạn trong giai đoạn này.",
    en: "We try our best to provide the best learning tools to the community at the smallest cost. We are working on our pricing. In the meantime, please contact us if you want to upgrade your account. We are happy to upgrade it for free for the time being.",
  },
  pricing_table_title: { vi: "Bảng so sánh các gói", en: "Tier Comparison" },
  pricing_row_lookups: { vi: "Lượt tra cứu tự động", en: "Auto-Lookups" },
  pricing_row_vocab: { vi: "Danh sách từ vựng", en: "Vocabulary Lists" },
  pricing_row_ads: { vi: "Quảng cáo", en: "Ads" },
  pricing_row_rate: { vi: "Giá", en: "Rate" },
  pricing_yes: { vi: "Có", en: "Yes" },
  pricing_no: { vi: "Không", en: "No" },
  pricing_tbd: { vi: "Sẽ cập nhật sau", en: "To be updated" },
  pricing_free_rate: { vi: "Miễn phí", en: "Free" },
  pricing_lookup_explainer_title: { vi: "Lượt tra cứu được tính như thế nào?", en: "What counts as a lookup?" },
  pricing_lookup_explainer_body: {
    vi: "Một lượt tra cứu chỉ bị trừ khi bạn tra một chữ Hán hoặc từ vựng MỚI mà hệ thống AI của chúng tôi chưa từng phân tích trước đó. Nếu chữ hoặc từ đó đã có sẵn trong kho dữ liệu công khai (do chúng tôi hoặc người dùng khác đã thêm vào), bạn sẽ nhận được kết quả ngay lập tức và hoàn toàn miễn phí — không bị trừ vào số lượt tra cứu của bạn.",
    en: "A lookup is only deducted when you look up a new character or word that our AI system hasn't already analyzed. If that character or word already exists in our shared public data (added by us or by another user), you get the result instantly and completely free — it does not use up any of your lookup allowance.",
  },
  pricing_hsk12: { vi: "New HSK 1 & 2", en: "New HSK 1 & 2" },
  pricing_hsk34: { vi: "New HSK 3 & 4", en: "New HSK 3 & 4" },
  pricing_hsk5: { vi: "New HSK 5", en: "New HSK 5" },
  pricing_free_vocab: { vi: "Danh sách cơ bản", en: "Basic lists" },
  pricing_courses_title: { vi: "Khóa học riêng", en: "Dedicated Courses" },
  pricing_courses_body: {
    vi: "Ngoài các gói thành viên, chúng tôi cũng tổ chức các khóa học với bộ từ vựng được biên soạn riêng cho từng khóa. Nếu bạn đang theo học một khóa cụ thể, tài khoản của bạn sẽ được cấp quyền truy cập vào danh sách từ vựng riêng của khóa đó.",
    en: "Beyond the standard tiers, we also run dedicated courses with vocabulary lists curated specifically for each one. If you're enrolled in a particular course, your account is granted access to that course's own vocabulary lists.",
  },
  footer_copyright: { vi: "Bản quyền © 2026 MinouQ", en: "Copyright © 2026 MinouQ" },
  blog_coming_soon_title: { vi: "Blog sắp ra mắt", en: "Blog Coming Soon" },
  blog_coming_soon_body: {
    vi: "Chúng tôi đang chuẩn bị các bài viết về mẹo học chữ Hán, bộ thủ, và phương pháp luyện viết. Quay lại sau nhé!",
    en: "We're preparing articles on Hanzi learning tips, radicals, and writing practice methods. Check back soon!",
  },
  blog_page_title: { vi: "Blog", en: "Blog" },
  blog_cat_all: { vi: "Tất cả", en: "All" },
  blog_cat_news: { vi: "Tin tức", en: "News" },
  blog_cat_resources: { vi: "Tài liệu miễn phí", en: "Free Resources" },
  blog_cat_founder: { vi: "Góc nhìn", en: "From the Founder" },
  blog_read_more: { vi: "Đọc tiếp →", en: "Read more →" },
  blog_collapse: { vi: "Thu gọn ↑", en: "Collapse ↑" },
  blog_external_link: { vi: "Đọc bài viết đầy đủ →", en: "Read the full article →" },
  blog_comments_title: (n) => ({ vi: `Bình luận (${n})`, en: `Comments (${n})` }),
  blog_comments_none: { vi: "Chưa có bình luận nào. Hãy là người đầu tiên!", en: "No comments yet. Be the first!" },
  blog_leave_comment: { vi: "Để lại bình luận", en: "Leave a comment" },
  blog_comment_name_placeholder: { vi: "Tên của bạn", en: "Your name" },
  blog_comment_placeholder: { vi: "Viết bình luận của bạn…", en: "Write your comment…" },
  blog_comment_moderation_note: {
    vi: "Bình luận của bạn sẽ hiển thị sau khi được quản trị viên duyệt.",
    en: "Your comment will appear after an admin approves it.",
  },
  blog_comment_submit: { vi: "Gửi bình luận", en: "Submit Comment" },
  blog_comment_thanks: { vi: "Cảm ơn bạn! Bình luận đang chờ duyệt.", en: "Thanks! Your comment is awaiting approval." },
  blog_empty: { vi: "Chưa có bài viết nào trong mục này.", en: "No posts in this category yet." },
  blog_read_time: (n) => ({ vi: `${n} phút đọc`, en: `${n} min read` }),
  admin_blog_title: { vi: "Quản lý bài viết Blog", en: "Blog Post Management" },
  admin_blog_new_post: { vi: "+ Bài viết mới", en: "+ New Post" },
  admin_blog_title_placeholder: { vi: "Tiêu đề bài viết", en: "Post title" },
  admin_blog_body_placeholder: { vi: "Nội dung bài viết…", en: "Post content…" },
  admin_blog_link_placeholder: { vi: "Đường dẫn ngoài (tùy chọn)", en: "External link (optional)" },
  admin_blog_published: { vi: "Đã đăng", en: "Published" },
  admin_blog_draft: { vi: "Bản nháp", en: "Draft" },
  admin_blog_save: { vi: "Lưu bài viết", en: "Save Post" },
  admin_blog_cancel: { vi: "Hủy", en: "Cancel" },
  admin_blog_edit: { vi: "Sửa", en: "Edit" },
  admin_blog_delete: { vi: "Xóa", en: "Delete" },
  admin_blog_publish_action: { vi: "Đăng bài", en: "Publish" },
  admin_blog_unpublish_action: { vi: "Gỡ đăng", en: "Unpublish" },
  admin_blog_confirm_delete: { vi: "Xóa bài viết này?", en: "Delete this post?" },
  admin_blog_need_title_body: { vi: "Vui lòng nhập tiêu đề và nội dung.", en: "Please enter a title and content." },
  admin_blog_none: { vi: "Chưa có bài viết nào.", en: "No posts yet." },
  about_title: { vi: "Về MinouQ", en: "About MinouQ" },
  about_body: {
    vi: "MinouQ Chinese là công cụ học chữ Hán được xây dựng để giúp người học hiểu chữ Hán qua cách phân tích các bộ thành phần, thay vì học thuộc lòng. Chúng tôi tin rằng việc hiểu cấu tạo của một chữ Hán sẽ giúp việc ghi nhớ trở nên tự nhiên và bền vững hơn.",
    en: "MinouQ Chinese is a Hanzi-learning tool built to help learners understand Chinese characters through their component structure, rather than by rote memorization. We believe that understanding how a character is built makes it easier to remember naturally and for the long term.",
  },
  about_visit_line: {
    vi: "Để tìm hiểu thêm về các bài viết, sản phẩm, và nghiên cứu khác của MinouQ, vui lòng truy cập chúng tôi tại",
    en: "To learn more about other articles, products, and research by MinouQ, please visit us at",
  },
  feedback_title: { vi: "Chúng tôi rất muốn nghe ý kiến của bạn", en: "We'd Love Your Feedback" },
  feedback_body: {
    vi: "Bạn thấy điều gì hữu ích? Điều gì cần cải thiện? Hãy cho chúng tôi biết bên dưới.",
    en: "What's working well? What could be better? Let us know below.",
  },
  feedback_contact_line: {
    vi: "Hoặc liên hệ trực tiếp với chúng tôi tại hello@minouq.com",
    en: "Or contact us directly at hello@minouq.com",
  },
  feedback_placeholder: { vi: "Viết góp ý của bạn ở đây…", en: "Write your feedback here…" },
  feedback_email_placeholder: { vi: "Email của bạn (không bắt buộc)", en: "Your email (optional)" },
  admin_feedback_title: { vi: "Góp ý từ người dùng", en: "User Feedback" },
  admin_feedback_none: { vi: "Chưa có góp ý nào.", en: "No feedback yet." },
  admin_feedback_no_email: { vi: "không có email", en: "no email" },
  admin_feedback_account: { vi: "Tài khoản:", en: "Account:" },
  admin_feedback_submitted_email: { vi: "Email đã nhập:", en: "Submitted email:" },
  admin_feedback_guest: { vi: "Khách (chưa đăng nhập)", en: "Guest (not signed in)" },

  // Admin panel - user management
  admin_nav_users: { vi: "Người dùng", en: "Users" },
  admin_nav_lists: { vi: "Danh sách", en: "Lists" },
  admin_nav_feedback: { vi: "Góp ý", en: "Feedback" },
  admin_nav_blog: { vi: "Blog", en: "Blog" },
  admin_nav_comments: { vi: "Bình luận", en: "Comments" },
  admin_comments_title: { vi: "Duyệt bình luận blog", en: "Blog Comment Moderation" },
  admin_comment_filter_pending: { vi: "Chờ duyệt", en: "Pending" },
  admin_comment_filter_approved: { vi: "Đã duyệt", en: "Approved" },
  admin_comment_filter_rejected: { vi: "Đã từ chối", en: "Rejected" },
  admin_comment_filter_all: { vi: "Tất cả", en: "All" },
  admin_comment_none: { vi: "Không có bình luận nào.", en: "No comments here." },
  admin_comment_on_post: { vi: "Trên bài viết:", en: "On post:" },
  admin_comment_unknown_post: { vi: "(không tìm thấy bài viết)", en: "(post not found)" },
  admin_comment_no_email: { vi: "không có email", en: "no email" },
  admin_comment_approve: { vi: "Duyệt", en: "Approve" },
  admin_comment_reject: { vi: "Từ chối", en: "Reject" },
  admin_comment_delete: { vi: "Xóa", en: "Delete" },
  admin_comment_confirm_delete: { vi: "Xóa bình luận này?", en: "Delete this comment?" },
  admin_nav_userlib: { vi: "Thư viện người dùng", en: "User Libraries" },
  admin_lib_search_placeholder: { vi: "Tìm người dùng theo email…", en: "Search for a user by email…" },
  admin_lib_viewing_banner: (email) => ({ vi: `Đang xem thư viện của ${email} — chỉ xem, không phải của bạn`, en: `Viewing ${email}'s library — read-only, not yours` }),
  admin_lib_exit: { vi: "✕ Thoát", en: "✕ Exit" },
  admin_lib_copy: { vi: "Sao chép", en: "Copy" },
  admin_lib_copy_list: { vi: "Sao chép cả danh sách", en: "Copy entire list" },
  admin_lib_copy_deck: { vi: "Sao chép bộ sưu tập", en: "Copy deck" },
  admin_lib_copy_done: { vi: "Đã sao chép vào thư viện của bạn.", en: "Copied to your library." },
  admin_lib_copy_error: { vi: "Không thể sao chép.", en: "Could not copy." },
  admin_lib_copy_list_done: (n, total) => ({ vi: `Đã sao chép ${n}/${total} mục.`, en: `Copied ${n}/${total} items.` }),
  admin_lib_copy_deck_done: (n) => ({ vi: `Đã sao chép bộ sưu tập cùng ${n} mục.`, en: `Copied the deck along with ${n} items.` }),
  admin_lib_no_content: { vi: "Người dùng này chưa có nội dung nào.", en: "This user hasn't added anything yet." },
  admin_lib_new_badge: { vi: "MỚI", en: "NEW" },
  admin_lib_only_new: { vi: "Chỉ hiện mục mới", en: "Show only new items" },
  admin_lib_chars_title: { vi: "Hán tự của họ", en: "Their Characters" },
  admin_lib_words_title: { vi: "Từ vựng của họ", en: "Their Words" },
  admin_lib_bushou_title: { vi: "Bộ thủ của họ", en: "Their Radicals" },
  admin_lib_decks_title: { vi: "Bộ sưu tập của họ", en: "Their Decks" },
  admin_lib_lists_in: { vi: "Danh sách:", en: "List:" },
  admin_lib_unlisted: { vi: "Chưa phân loại", en: "Not in any list" },
  admin_nav_suggestions: { vi: "Đề xuất", en: "Suggestions" },
  admin_suggestions_title: { vi: "Đề xuất chỉnh sửa từ người dùng", en: "User Revision Suggestions" },
  admin_suggestion_filter_all: { vi: "Tất cả", en: "All" },
  admin_suggestion_status_new: { vi: "Mới", en: "New" },
  admin_suggestion_status_read: { vi: "Đã đọc", en: "Read" },
  admin_suggestion_status_revised: { vi: "Đã sửa theo đề xuất", en: "Revised as Suggested" },
  admin_suggestion_status_ignored: { vi: "Bỏ qua", en: "Ignored" },
  admin_suggestion_none: { vi: "Không có đề xuất nào.", en: "No suggestions here." },
  admin_suggestion_unknown_item: { vi: "(không tìm thấy mục này)", en: "(item not found)" },
  admin_suggestion_no_email: { vi: "không có email", en: "no email" },
  admin_suggestion_type_char: { vi: "Hán tự", en: "Character" },
  admin_suggestion_type_word: { vi: "Từ vựng", en: "Word" },
  admin_suggestion_type_bushou: { vi: "Bộ thủ", en: "Radical" },
  admin_suggestion_confirm_delete: { vi: "Xóa đề xuất này?", en: "Delete this suggestion?" },
  admin_suggestion_mark_read: { vi: "Đánh dấu đã đọc", en: "Mark Read" },
  admin_suggestion_mark_revised: { vi: "Đã sửa theo đề xuất", en: "Mark Revised" },
  admin_suggestion_mark_ignored: { vi: "Bỏ qua", en: "Ignore" },
  admin_nav_decks: { vi: "Bộ sưu tập", en: "Decks" },
  admin_deck_title: { vi: "Quản lý bộ sưu tập", en: "Deck Management" },
  admin_deck_new: { vi: "+ Bộ sưu tập mới", en: "+ New Deck" },
  admin_deck_name_placeholder: { vi: "Tên bộ sưu tập", en: "Deck name" },
  admin_deck_description_placeholder: { vi: "Mô tả (tùy chọn)", en: "Description (optional)" },
  admin_deck_included_lists: { vi: "Danh sách được gộp vào:", en: "Lists included:" },
  admin_deck_none_included: { vi: "Chưa có danh sách nào.", en: "No lists added yet." },
  admin_deck_type_char: { vi: "Hán tự", en: "Characters" },
  admin_deck_type_word: { vi: "Từ vựng", en: "Words" },
  admin_deck_type_bushou: { vi: "Bộ thủ", en: "Radicals" },
  admin_deck_choose_list: { vi: "— Chọn danh sách —", en: "— Choose a list —" },
  admin_deck_add: { vi: "+ Thêm", en: "+ Add" },
  admin_deck_save: { vi: "Lưu bộ sưu tập", en: "Save Deck" },
  admin_deck_cancel: { vi: "Hủy", en: "Cancel" },
  admin_deck_edit: { vi: "Sửa", en: "Edit" },
  admin_deck_delete: { vi: "Xóa", en: "Delete" },
  admin_deck_none: { vi: "Chưa có bộ sưu tập nào.", en: "No decks yet." },
  admin_deck_need_name: { vi: "Vui lòng nhập tên bộ sưu tập.", en: "Please enter a deck name." },
  admin_deck_confirm_delete: { vi: "Xóa bộ sưu tập này? Các danh sách bên trong sẽ không bị xóa.", en: "Delete this deck? The lists inside it won't be deleted." },
  admin_deck_list_count: (n) => ({ vi: `${n} danh sách`, en: `${n} list${n === 1 ? "" : "s"}` }),
  admin_users_title: { vi: "Quản trị người dùng", en: "User Management" },
  admin_search_email_placeholder: { vi: "Tìm theo email…", en: "Search by email…" },
  admin_all_tiers: { vi: "Tất cả gói", en: "All Tiers" },
  admin_all_status: { vi: "Tất cả trạng thái", en: "All Statuses" },
  admin_status_enabled: { vi: "Đang hoạt động", en: "Active" },
  admin_status_disabled: { vi: "Đã vô hiệu hóa", en: "Disabled" },
  admin_all_courses: { vi: "Tất cả khóa học", en: "All Courses" },
  admin_refresh: { vi: "⟳ Làm mới", en: "⟳ Refresh" },
  admin_user_count: (filtered, total) => ({ vi: `${filtered} / ${total} người dùng`, en: `${filtered} / ${total} users` }),
  admin_invalid_limit: { vi: "Giới hạn không hợp lệ.", en: "Invalid limit." },
  admin_need_course_name: { vi: "Vui lòng nhập tên khóa học.", en: "Please enter a course name." },
  admin_save_failed: (msg) => ({ vi: `Không thể lưu: ${msg}`, en: `Could not save: ${msg}` }),
  admin_saved: { vi: "Đã lưu.", en: "Saved." },
  admin_confirm_reset_usage: { vi: "Đặt lại lượt tra cứu về 0 cho người dùng này?", en: "Reset lookup count to 0 for this user?" },
  admin_reset_failed: (msg) => ({ vi: `Không thể đặt lại: ${msg}`, en: `Could not reset: ${msg}` }),
  admin_reset_done: { vi: "Đã đặt lại.", en: "Reset." },
  admin_confirm_disable: (name) => ({
    vi: `Vô hiệu hóa tài khoản "${name}"? Người này sẽ không thể tra cứu tự động cho đến khi được kích hoạt lại.`,
    en: `Disable account "${name}"? This person won't be able to use auto-lookup until re-enabled.`,
  }),
  admin_update_failed: (msg) => ({ vi: `Không thể cập nhật: ${msg}`, en: `Could not update: ${msg}` }),
  admin_disabled_done: { vi: "Đã vô hiệu hóa.", en: "Disabled." },
  admin_enabled_done: { vi: "Đã kích hoạt lại.", en: "Re-enabled." },
  admin_course_name_placeholder: { vi: "Tên khóa học…", en: "Course name…" },
  admin_save: { vi: "Lưu", en: "Save" },
  admin_cancel: { vi: "Hủy", en: "Cancel" },
  admin_disabled_badge: { vi: "Đã vô hiệu hóa", en: "Disabled" },
  admin_edit: { vi: "Sửa", en: "Edit" },
  admin_reset_to_zero: { vi: "Đặt lại về 0", en: "Reset to 0" },
  admin_reenable: { vi: "✓ Kích hoạt lại", en: "✓ Re-enable" },
  admin_disable: { vi: "🚫 Vô hiệu hóa", en: "🚫 Disable" },
  admin_no_users_found: { vi: "Không tìm thấy người dùng.", en: "No users found." },

  // Admin panel - list management
  admin_list_mgmt_title: { vi: "Quản lý danh sách", en: "List Management" },
  admin_list_mgmt_description: {
    vi: 'Danh sách chưa cấu hình bên dưới mặc định mở cho mọi người. "Chỉ admin" ẩn hoàn toàn khỏi người dùng thường. Chọn gói và/hoặc gán khóa học cụ thể để giới hạn quyền xem nội dung (tên danh sách vẫn hiển thị cho mọi người, trừ khi chọn "Chỉ admin").',
    en: 'Lists not configured below default to open for everyone. "Admin only" hides it completely from regular users. Choose tiers and/or grant specific courses to restrict content access (the list name is still visible to everyone, unless "Admin only" is selected).',
  },
  admin_admin_only_checkbox: { vi: "Chỉ admin", en: "Admin only" },
  admin_allowed_tiers_label: { vi: "Gói được phép xem:", en: "Tiers allowed to view:" },
  admin_course_grants_label: {
    vi: 'Cấp riêng cho khóa học cụ thể (dành cho gói "Enrolled Course"):',
    en: 'Grant specific courses (for the "Enrolled Course" tier):',
  },
  admin_add: { vi: "+ Thêm", en: "+ Add" },
  admin_open_to_all: { vi: "Mở cho tất cả", en: "Open to all" },
  admin_admin_only_badge: { vi: "🔒 Chỉ admin", en: "🔒 Admin only" },
  admin_course_count: (n) => ({ vi: `${n} khóa học`, en: `${n} course${n === 1 ? "" : "s"}` }),
  admin_no_one_can_view: { vi: "Không ai được xem", en: "No one can view" },
  admin_no_lists_yet: { vi: "Chưa có danh sách nào.", en: "No lists yet." },
  admin_feedback_delete: { vi: "Xóa", en: "Delete" },
  admin_feedback_mark_read: { vi: "Đánh dấu đã đọc", en: "Mark Read" },
  feedback_submit: { vi: "Gửi góp ý", en: "Submit Feedback" },
  feedback_sending: { vi: "Đang gửi…", en: "Sending…" },
  feedback_success: { vi: "Cảm ơn bạn đã góp ý!", en: "Thank you for your feedback!" },
  feedback_error: { vi: "Gửi thất bại. Vui lòng thử lại.", en: "Failed to send. Please try again." },
  feedback_empty: { vi: "Vui lòng nhập góp ý trước khi gửi.", en: "Please write something before submitting." },
  tab_admin: { vi: "⚙ Quản trị", en: "⚙ Admin" },

  // Header
  header_subtitle: {
    vi: "Học chữ Hán qua cách phân tích các bộ thành phần và luyện viết tay trên máy",
    en: "Study Chinese characters through understanding radical components and practicing handwriting on smart device",
  },

  // Ghép bộ thủ (Play tab)
  play_all_lists: { vi: "Tất cả danh sách", en: "All lists" },
  play_review_list: (n) => ({ vi: `🔁 Cần ôn lại (${n})`, en: `🔁 Needs Review (${n})` }),
  play_difficulty_label: { vi: "Độ khó:", en: "Difficulty:" },
  play_start: { vi: "Bắt đầu", en: "Start" },
  play_pinyin: { vi: "Pinyin:", en: "Pinyin:" },
  play_han_viet: { vi: "Hán Việt:", en: "Sino-Vietnamese:" },
  play_wrong_answer: { vi: "✗ Chưa đúng. Đáp án đúng:", en: "✗ Not quite. Correct answer:" },
  play_answer_reveal: { vi: "💡 Đáp án:", en: "💡 Answer:" },
  play_correct_prefix: { vi: "✓ Chính xác!", en: "✓ Correct!" },
  play_score: { vi: "Điểm:", en: "Score:" },
  play_streak: { vi: "Chuỗi đúng:", en: "Streak:" },
  play_playable_count: (n, isWord) => ({
    vi: `${n} ${isWord ? "mục" : "chữ"} có thể học`,
    en: `${n} ${isWord ? "item" : "character"}${n === 1 ? "" : "s"} available`,
  }),
  play_needs_review_suffix: { vi: "(🔁 Cần ôn lại)", en: "(🔁 Needs Review)" },
  play_undo: { vi: "Undo - Chọn lại", en: "Undo - Reselect" },
  play_show_answer: { vi: "💡 Xem đáp án", en: "💡 Show Answer" },
  play_next_char: { vi: "Chữ tiếp theo →", en: "Next Character →" },
  play_no_playable_review: {
    vi: 'Danh sách ôn lại đang trống — nó chỉ chứa những chữ bạn đã dùng nút "Xem đáp án". Trả lời đúng một chữ sẽ tự động xóa nó khỏi danh sách này.',
    en: 'Your review list is empty — it only holds characters you\'ve used "Show Answer" on. Answering one correctly automatically removes it from this list.',
  },
  play_no_playable_list: {
    vi: 'Danh sách này chưa có chữ nào chơi được — có thể vì các chữ trong đó chưa có bộ thủ cấu thành. Ở tab "Tạo thẻ từ mới", hãy dùng nút "🔍 Tự động điền" trước khi lưu để hệ thống tự nhận diện bộ thủ.',
    en: 'This list has no playable characters yet — likely because they don\'t have components assigned. In "Create New Card", use the "🔍 Auto-fill" button before saving so the system can detect the components.',
  },
  play_no_data: {
    vi: 'Chưa có chữ nào trong kho dữ liệu. Hãy thêm chữ ở tab "Tạo thẻ từ mới".',
    en: 'No characters in storage yet. Add some in the "Create New Card" tab.',
  },

  // Flashcard
  fc_title: { vi: "Ôn tập bằng thẻ ghi nhớ", en: "Review with Flashcards" },
  fc_all_lists: { vi: "Tất cả danh sách", en: "All lists" },
  fc_content_words: { vi: "Chữ & Từ", en: "Characters & Words" },
  fc_content_radicals: { vi: "Bộ thủ", en: "Radicals" },
  fc_content_deck: { vi: "Bộ sưu tập", en: "Deck" },
  fc_due_today: (n) => ({ vi: `${n} thẻ cần ôn hôm nay`, en: `${n} cards due today` }),
  fc_start: { vi: "Bắt đầu", en: "Start" },
  fc_progress: (reviewed, remaining) => ({
    vi: `Còn ${remaining} thẻ · đã ôn ${reviewed}`,
    en: `${remaining} cards left · ${reviewed} reviewed`,
  }),
  fc_flip: { vi: "Lật thẻ", en: "Flip card" },
  fc_again: { vi: "Chưa nhớ", en: "Again" },
  fc_hard: { vi: "Khó", en: "Hard" },
  fc_good: { vi: "Bình thường", en: "Good" },
  fc_easy: { vi: "Dễ", en: "Easy" },
  fc_end: { vi: "Kết thúc", en: "End session" },
  fc_complete_title: { vi: "Hoàn thành! 🎉", en: "Complete! 🎉" },
  fc_complete_summary: (reviewed, again) => ({
    vi: `Đã ôn ${reviewed} thẻ, ${again} thẻ cần ôn lại sớm.`,
    en: `Reviewed ${reviewed} cards, ${again} need review again soon.`,
  }),
  fc_done: { vi: "Xong", en: "Done" },

  // Luyện viết (Writing practice)
  wp_title: { vi: "Luyện viết theo nét", en: "Guided Writing Practice" },
  wp_search_placeholder: { vi: "Tìm một chữ cụ thể để luyện…", en: "Search for a character to practice…" },
  wp_or_by_list: { vi: "— hoặc luyện theo danh sách —", en: "— or practice by list —" },
  wp_all_lists: { vi: "Tất cả danh sách", en: "All lists" },
  wp_description: {
    vi: "Xem thứ tự nét trước, sau đó tự viết theo, dùng chuột (máy tính) hoặc ngón tay (màn hình cảm ứng).",
    en: "Watch the stroke order first, then write it yourself, using a mouse (computer) or your finger (touchscreen).",
  },
  wp_start: { vi: "Bắt đầu", en: "Start" },
  wp_replay: { vi: "↻ Xem lại", en: "↻ Replay" },
  wp_begin_writing: { vi: "✍️ Bắt đầu viết", en: "✍️ Start Writing" },
  wp_retrace: { vi: "↻ Tô lại", en: "↻ Trace Again" },
  wp_connect_dots: { vi: "🔵 Nối điểm →", en: "🔵 Connect the Dots →" },
  wp_skip_step: { vi: "Bỏ qua bước này →", en: "Skip this step →" },
  wp_choose_another: { vi: "🔍 Chọn chữ khác", en: "🔍 Choose Another" },
  wp_next_char: { vi: "Chữ tiếp theo →", en: "Next Character →" },
  wp_write_from_memory: { vi: "✏️ Viết từ trí nhớ →", en: "✏️ Write From Memory →" },
  wp_end: { vi: "Kết thúc", en: "End" },
  wp_reveal: { vi: "Hiện chữ đúng", en: "Reveal" },
  wp_undo_stroke: { vi: "↩ Xóa nét trước", en: "↩ Undo Last Stroke" },
  wp_recall_erase: { vi: "↩ Tẩy nét cuối", en: "↩ Erase" },
  wp_clear_all: { vi: "🗑 Xóa hết, viết lại", en: "🗑 Start Over" },
  wp_brush_size: { vi: "Cỡ bút:", en: "Brush size:" },
  wp_brush_thin: { vi: "Mảnh", en: "Thin" },
  wp_brush_normal: { vi: "Vừa", en: "Normal" },
  wp_brush_thick: { vi: "Đậm", en: "Thick" },
  wp_no_stroke_data: (char) => ({
    vi: `Chưa có dữ liệu nét bút cho chữ "${char}" — bấm "Chữ tiếp theo" để bỏ qua.`,
    en: `No stroke data available for "${char}" — click "Next Character" to skip.`,
  }),
  wp_complete_no_mistakes: { vi: "✓ Hoàn thành! Không có lỗi nào.", en: "✓ Complete! No mistakes." },
  wp_complete_with_mistakes: (n) => ({ vi: `✓ Hoàn thành! ${n} lỗi.`, en: `✓ Complete! ${n} mistake(s).` }),
  wp_dots_connect_prefix: { vi: "nối điểm", en: "connect the" },
  wp_dots_green: { vi: "xanh dương", en: "blue" },
  wp_dots_start_label: { vi: "(bắt đầu)", en: "(start)" },
  wp_dots_to: { vi: "tới điểm", en: "dot to the" },
  wp_dots_red: { vi: "đỏ", en: "red" },
  wp_dots_end_label: { vi: "(kết thúc)", en: "(end) dot" },
  wp_dots_stroke_progress: (i, total) => ({ vi: `Nét ${i} / ${total} —`, en: `Stroke ${i} / ${total} —` }),
  wp_dots_correct: { vi: "✓ Đúng!", en: "✓ Correct!" },
  wp_dots_wrong: { vi: "Chưa đúng, thử lại", en: "Not quite, try again" },
  wp_dots_all_done: (n) => ({ vi: `✓ Đã nối xong tất cả ${n} nét!`, en: `✓ Connected all ${n} strokes!` }),
  wp_dots_load_error: (char) => ({
    vi: `Không tải được dữ liệu nét cho chữ "${char}" — bấm "Bỏ qua bước này" để tiếp tục.`,
    en: `Could not load stroke data for "${char}" — click "Skip this step" to continue.`,
  }),
  wp_prev_page: { vi: "← Trang trước", en: "← Previous Page" },
  wp_next_page: { vi: "Trang sau →", en: "Next Page →" },

  // Tạo thẻ từ mới (Add tab)
  add_intro: {
    vi: "Nhập một chữ Hán hoàn chỉnh cùng nghĩa, pinyin, và một danh sách. Nếu chữ đã được phân tích trong kho dữ liệu chung, lượt tra cứu của bạn sẽ không bị tính. Lưu ý: ở đây bạn chỉ có thể gán một danh sách cho thẻ mới — để tạo bộ sưu tập (deck), vui lòng vào mục Quản lý → Quản lý thư viện.",
    en: "Enter a complete Chinese character along with its meaning, pinyin, and a list. If a character has been broken down in shared storage, your lookup quota will not be counted. Note: here you can only assign a list to a new card — to create a deck, go to Management → Library Management.",
  },
  add_char_label: { vi: "Chữ Hán hoàn chỉnh", en: "Complete Chinese Character" },
  add_meaning_en_label: { vi: "Nghĩa (English)", en: "Meaning (English)" },
  add_meaning_vi_label: { vi: "Nghĩa (Tiếng Việt)", en: "Meaning (Vietnamese)" },
  add_pinyin_label: { vi: "Pinyin", en: "Pinyin" },
  add_hanviet_label: { vi: "Âm Hán Việt (Sino-Vietnamese)", en: "Sino-Vietnamese Reading" },
  add_lists_label: { vi: "Danh sách (Lists)", en: "Lists" },
  add_autofill: { vi: "🔍 Tự động điền", en: "🔍 Auto-fill" },
  add_looking_up: { vi: "Đang tra…", en: "Looking up…" },
  add_save: { vi: "Lưu lại", en: "Save" },
  add_note_button: { vi: "⚠️ Lưu ý", en: "⚠️ Note" },
  add_fill_required: {
    vi: "Vui lòng điền đầy đủ: chữ Hán, nghĩa, pinyin, âm Hán Việt.",
    en: "Please fill in all fields: character, meaning, pinyin, Sino-Vietnamese reading.",
  },
  add_char_exists: (char) => ({
    vi: `Chữ "${char}" đã có trong kho dữ liệu.`,
    en: `Character "${char}" already exists in storage.`,
  }),
  add_no_components_note: {
    vi: ' (chưa có bộ thủ cấu thành nên sẽ không xuất hiện trong chế độ Chơi — dùng nút "🔍 Tự động điền" hoặc thêm bộ thủ thủ công trước khi lưu)',
    en: ' (no components assigned yet, so it won\'t appear in Play mode — use "🔍 Auto-fill" or add components manually before saving)',
  },
  add_char_success: (char, lists, note) => ({
    vi: `Đã thêm chữ "${char}" vào danh sách "${lists}"!${note}`,
    en: `Added character "${char}" to list "${lists}"!${note}`,
  }),
  add_char_error: (msg) => ({
    vi: `Có lỗi xảy ra: ${msg}. Vui lòng thử lại.`,
    en: `An error occurred: ${msg}. Please try again.`,
  }),
  add_unknown_error: { vi: "không rõ nguyên nhân", en: "unknown cause" },

  // Bulk import panel
  bulk_toggle: (max) => ({
    vi: `Nhập hàng loạt (tối đa ${max} mục, chữ đơn hoặc từ nhiều chữ)`,
    en: `Bulk Import (up to ${max} items, single characters or multi-character words)`,
  }),
  bulk_instructions: (max) => ({
    vi: `Dán tối đa ${max} mục, mỗi mục một dòng. Một chữ đơn (vd: 好) sẽ được thêm như một chữ; hai chữ trở lên trên cùng một dòng (vd: 你好) sẽ được thêm như một từ. Mỗi mục chưa có sẽ được tra cứu tự động; mục đã có sẽ chỉ được gắn thêm tên danh sách này.`,
    en: `Paste up to ${max} items, one per line. A single character (e.g. 好) is added as a character; two or more characters on the same line (e.g. 你好) are added as a word. New items are looked up automatically; existing ones just get this list name added.`,
  }),
  bulk_count: (n, max) => ({ vi: `${n} / ${max} mục`, en: `${n} / ${max} items` }),
  bulk_lists_label: { vi: "Danh sách (có thể chọn nhiều)", en: "Lists (multiple allowed)" },
  bulk_list_placeholder: { vi: "vd: HSK1… rồi Enter", en: "e.g. HSK1… then Enter" },
  bulk_add: { vi: "+ Thêm", en: "+ Add" },
  bulk_start: { vi: "Bắt đầu nhập", en: "Start Import" },
  bulk_stop: { vi: "Dừng lại", en: "Stop" },
  bulk_reset: { vi: "Làm mới", en: "Reset" },
  bulk_processing: (item) => ({ vi: `— đang xử lý: ${item}`, en: `— processing: ${item}` }),
  bulk_complete: { vi: "— hoàn tất", en: "— complete" },
  bulk_summary: (added, tagged) => ({
    vi: `✓ Đã thêm mới ${added} mục, gắn thêm danh sách cho ${tagged} mục đã có sẵn.`,
    en: `✓ Added ${added} new items, tagged ${tagged} existing items with this list.`,
  }),
  bulk_errors: (n, items) => ({ vi: `✗ ${n} mục thất bại: ${items}`, en: `✗ ${n} item(s) failed: ${items}` }),
  bulk_retry_hint: {
    vi: 'Có thể thử lại bằng cách dán riêng các mục này và bấm "Bắt đầu nhập" lần nữa.',
    en: 'You can retry by pasting just these items and clicking "Start Import" again.',
  },
  bulk_no_chars_found: { vi: "Không tìm thấy chữ Hán nào trong ô dán.", en: "No Chinese characters found in the pasted text." },
  bulk_need_list: { vi: "Vui lòng chọn hoặc thêm ít nhất một danh sách.", en: "Please select or add at least one list." },

  // Add-word panel (multi-character word)
  word_toggle: { vi: "Thêm từ nhiều chữ (vd: 你好)", en: "Add Multi-Character Word (e.g. 你好)" },
  word_instructions: {
    vi: "Gõ một từ có từ 2 chữ Hán trở lên, rồi rời khỏi ô (hoặc bấm nút) để tự động điền pinyin, nghĩa, Hán Việt, và bộ thủ cho từng chữ còn thiếu — tất cả trong một bước.",
    en: "Type a word with 2 or more Chinese characters, then leave the field (or click the button) to auto-fill pinyin, meaning, Sino-Vietnamese, and components for each missing character — all in one step.",
  },
  word_looking_up: { vi: "Đang tra…", en: "Looking up…" },
  word_autofill: { vi: "🔍 Tự động điền", en: "🔍 Auto-fill" },
  word_ready: { vi: "✓ sẵn sàng", en: "✓ ready" },
  word_no_components: { vi: "chưa có bộ thủ", en: "no components" },
  word_failed: { vi: "thất bại", en: "failed" },
  word_pinyin_label: { vi: "Pinyin cả từ", en: "Whole-word Pinyin" },
  word_meaning_en_label: { vi: "Nghĩa cả từ (English)", en: "Whole-word Meaning (English)" },
  word_meaning_vi_label: { vi: "Nghĩa cả từ (Tiếng Việt)", en: "Whole-word Meaning (Vietnamese)" },
  word_hanviet_label: { vi: "Hán Việt", en: "Sino-Vietnamese" },
  word_lists_label: { vi: "Danh sách (Lists)", en: "Lists" },
  word_auto_or_manual: { vi: "tự động điền, hoặc nhập tay", en: "auto-filled, or type manually" },
  word_list_placeholder: { vi: "vd: Thành ngữ… rồi Enter", en: "e.g. Idioms… then Enter" },
  word_add_list: { vi: "+ Thêm", en: "+ Add" },
  word_save: { vi: "Lưu từ", en: "Save Word" },
  word_lookup_failed: {
    vi: "Tra cứu từ thất bại. Vui lòng nhập tay pinyin / nghĩa / Hán Việt.",
    en: "Word lookup failed. Please enter pinyin / meaning / Sino-Vietnamese manually.",
  },
  word_need_2_chars: { vi: "Từ cần có ít nhất 2 chữ Hán.", en: "The word needs at least 2 Chinese characters." },
  word_need_components: {
    vi: 'Mỗi chữ trong từ cần có bộ thủ trước — dùng nút "Tự động điền" bên dưới cho những chữ còn thiếu.',
    en: 'Every character in the word needs components first — use the "Auto-fill" button below for any that are missing.',
  },
  word_need_pinyin_meaning: { vi: "Vui lòng điền pinyin và nghĩa của từ.", en: "Please fill in the word's pinyin and meaning." },
  word_exists: (word) => ({ vi: `Từ "${word}" đã có trong kho dữ liệu.`, en: `Word "${word}" already exists in storage.` }),
  word_added_success: (word) => ({ vi: `Đã thêm từ "${word}"!`, en: `Added word "${word}"!` }),

  // Rename-list panel
  rename_toggle: { vi: "Đổi tên danh sách", en: "Rename List" },
  rename_description: {
    vi: "Đổi tên áp dụng cho mọi chữ và từ đang mang tên danh sách này — cả chữ/từ có sẵn lẫn chữ/từ của bạn.",
    en: "Renaming applies to every character and word currently tagged with this list — both built-in and your own.",
  },
  rename_choose_list: { vi: "— Chọn danh sách —", en: "— Choose a list —" },
  rename_new_name_placeholder: { vi: "Tên mới", en: "New name" },
  rename_renaming: { vi: "Đang đổi…", en: "Renaming…" },
  rename_button: { vi: "Đổi tên", en: "Rename" },
  rename_need_old: { vi: "Vui lòng chọn danh sách cần đổi tên.", en: "Please choose a list to rename." },
  rename_need_new: { vi: "Vui lòng nhập tên mới.", en: "Please enter a new name." },
  rename_same_name: { vi: "Tên mới trùng với tên cũ.", en: "The new name is the same as the old one." },
  rename_none_found: (name) => ({
    vi: `Không có chữ hoặc từ nào thuộc danh sách "${name}".`,
    en: `No characters or words belong to the list "${name}".`,
  }),
  rename_success: (oldName, newName, total) => ({
    vi: `Đã đổi tên "${oldName}" thành "${newName}" cho ${total} mục.`,
    en: `Renamed "${oldName}" to "${newName}" for ${total} item(s).`,
  }),

  // Additional Create New Card descriptions
  add_char_field_description: {
    vi: "Gõ chữ Hán rồi rời khỏi ô để tự động điền pinyin, nghĩa, Hán Việt, và bộ thủ cấu thành — bạn vẫn có thể sửa lại thủ công.",
    en: "Type a Chinese character then leave the field to auto-fill pinyin, meaning, Sino-Vietnamese, and components — you can still edit everything manually afterward.",
  },
  add_components_label: { vi: "Bộ thủ cấu thành (theo thứ tự)", en: "Components (in order)" },
  add_components_description: {
    vi: "Tự động điền ở trên sẽ gợi ý sẵn, hoặc bạn có thể thêm / xóa bộ thủ thủ công bên dưới. Cần bộ thủ để chữ này xuất hiện được ở chế độ Chơi.",
    en: "Auto-fill above suggests these automatically, or you can add / remove components manually below. Components are required for this character to appear in Play mode.",
  },
  add_component_input_placeholder: { vi: "gõ 1 bộ thủ, vd: 女", en: "type a component, e.g. 女" },
  add_component_add_button: { vi: "+ Thêm bộ thủ", en: "+ Add Component" },
  add_new_component_before: { vi: "Bộ thủ", en: "Component" },
  add_new_component_after: { vi: "chưa có trong kho — hãy điền thông tin:", en: "isn't in storage yet — please fill in its details:" },
  add_new_comp_pinyin: { vi: "pinyin", en: "pinyin" },
  add_new_comp_meaning: { vi: "meaning", en: "meaning" },
  add_new_comp_sv: { vi: "âm Hán Việt", en: "Sino-Vietnamese reading" },
  add_new_comp_strokes: { vi: "số nét (tùy chọn)", en: "stroke count (optional)" },
  add_confirm: { vi: "Xác nhận", en: "Confirm" },
  add_lists_description: {
    vi: 'Một chữ có thể thuộc nhiều danh sách cùng lúc — gõ tên rồi nhấn Enter hoặc "+ Thêm" để thêm từng danh sách.',
    en: 'A character can belong to multiple lists at once — type a name and press Enter or "+ Add" to add each list.',
  },
  add_autofill_failed: {
    vi: "Tra cứu tự động thất bại. Vui lòng nhập pinyin / nghĩa / Hán Việt thủ công.",
    en: "Auto-lookup failed. Please enter pinyin / meaning / Sino-Vietnamese manually.",
  },
  add_need_meaning_en: { vi: "Vui lòng điền Nghĩa (English) hoặc bỏ tick ô này.", en: "Please fill in Meaning (English), or untick this field." },
  add_need_meaning_vi: { vi: "Vui lòng điền Nghĩa (Tiếng Việt) hoặc bỏ tick ô này.", en: "Please fill in Meaning (Vietnamese), or untick this field." },
  add_need_sv: { vi: "Vui lòng điền Âm Hán Việt hoặc bỏ tick ô này.", en: "Please fill in Sino-Vietnamese, or untick this field." },
  add_need_list: { vi: "Vui lòng chọn hoặc thêm ít nhất một danh sách.", en: "Please select or add at least one list." },
  add_field_required_checkbox: { vi: "Cần trường này", en: "Required" },
  bulk_missing_pinyin: { vi: "không tra được pinyin", en: "pinyin not found" },
  bulk_missing_meaning_en: { vi: "thiếu Nghĩa (English)", en: "missing Meaning (English)" },
  bulk_missing_meaning_vi: { vi: "thiếu Nghĩa (Tiếng Việt)", en: "missing Meaning (Vietnamese)" },
  bulk_missing_sv: { vi: "thiếu Âm Hán Việt", en: "missing Sino-Vietnamese" },
  bulk_required_fields_label: { vi: "Trường bắt buộc:", en: "Required fields:" },
  bulk_field_meaning_en: { vi: "Nghĩa (English)", en: "Meaning (English)" },
  bulk_field_meaning_vi: { vi: "Nghĩa (Tiếng Việt)", en: "Meaning (Vietnamese)" },
  bulk_field_sv: { vi: "Âm Hán Việt", en: "Sino-Vietnamese" },
  add_list_placeholder: { vi: "vd: HSK1, Gia đình, Bài 5… rồi Enter", en: "e.g. HSK1, Family, Lesson 5… then Enter" },
  add_list_add_button: { vi: "+ Thêm", en: "+ Add" },
  quota_admin_usage: (n) => ({ vi: `${n} lượt đã dùng · không giới hạn`, en: `${n} lookups used · unlimited` }),
  quota_remaining: (remaining, limit) => ({
    vi: `${remaining}/${limit} lượt tra cứu còn lại`,
    en: `${remaining}/${limit} lookups remaining`,
  }),

  add_note_p1: {
    vi: "Lưu ý, thanh tra cứu đôi khi sẽ tách các bộ thành phần của chữ Hán chưa chính xác hoặc khác với nhu cầu của người học.",
    en: "Note: the lookup tool sometimes splits a character's components imprecisely or differently than a learner would want.",
  },
  add_note_p2: {
    vi: "Ví dụ: 超 sẽ được công cụ tách thành 走 và 召. Tuy nhiên, người học cũng có thể tách chữ này thành 3 bộ 走, 刀, 口, hoặc 走 và 召, tùy theo nguyện vọng cá nhân.",
    en: "Example: the tool would split 超 into 走 and 召. However, a learner could also split this character into 3 parts (走, 刀, 口) or into 走 and 召, depending on personal preference.",
  },
  add_note_p3: {
    vi: "Người học sẽ phải nhập thủ công các bộ thành phần trong trường hợp người học muốn tách bộ thành phần theo cách khác với mặc định của thanh công cụ.",
    en: "Learners will need to enter components manually if they want to split them differently from the tool's default.",
  },
  add_note_p4: {
    vi: "Ngoài ra, thanh công cụ đôi lúc vẫn có thể tách sai bộ thành phần. Người học cần tra soát lại với các hệ thống từ điển và nhập lại thủ công nếu phát hiện sai sót. Các từ điển tham khảo:",
    en: "Also, the tool can occasionally split components incorrectly. Double-check against a dictionary and re-enter manually if you spot an error. Reference dictionaries:",
  },

  // Bộ thủ (Radicals tab)
  radicals_header_p1: {
    vi: "Kho lưu trữ bao gồm các bộ thủ chính và các bộ thành phần cấu tạo chữ Hán. Các bộ thành phần này sẽ bao gồm các bộ thủ chính. Người học có thể tùy ý chỉnh sửa cấu tạo thành phần của mỗi chữ Hán tùy theo thói quen học của mỗi cá nhân.",
    en: "This library includes the main radicals as well as the component parts that make up Chinese characters. These components in turn include the main radicals. Learners can freely edit how each character's components are structured to match their own study habits.",
  },
  radicals_header_p2: {
    vi: "Ví dụ: 语 có thể tách thành 讠 và 吾, hoặc có thể tách 讠, 五, 口 tùy lựa chọn của người học.",
    en: "Example: 语 can be split into 讠 and 吾, or into 讠, 五, 口 — whichever the learner prefers.",
  },
  radicals_header_p3: {
    vi: "Lưu ý, thanh tra cứu đôi khi sẽ tách các bộ thành phần của chữ Hán chưa chính xác hoặc khác với nhu cầu của người học (như ví dụ phía trên). Người học cần tra soát lại với các hệ thống từ điển và nhập lại thủ công nếu phát hiện sai sót.",
    en: "Note: the lookup tool sometimes splits a character's components imprecisely or differently than a learner would want (as in the example above). Double-check against a dictionary and re-enter manually if you spot an error.",
  },
  radicals_search_placeholder: {
    vi: "Tìm bộ thủ theo chữ, pinyin, nghĩa, hoặc Hán Việt…",
    en: "Search radicals by character, pinyin, meaning, or Sino-Vietnamese…",
  },
  radicals_count: (filtered, total) => ({
    vi: `${filtered} / ${total} bộ thủ và bộ thành phần · sắp xếp theo số nét · bấm ✎ để sửa`,
    en: `${filtered} / ${total} radicals and components · sorted by stroke count · click ✎ to edit`,
  }),
  radicals_view_stroke_order: { vi: "Xem thứ tự nét bút", en: "View stroke order" },
  radicals_set_default_tooltip: { vi: "Đặt/cập nhật làm dữ liệu mặc định cho mọi người dùng mới", en: "Set/update as default data for all new users" },
  radicals_withdraw_tooltip: { vi: "Bấm để gỡ khỏi dữ liệu mặc định", en: "Click to remove from default data" },
  radicals_working: { vi: "Đang xử lý…", en: "Working…" },
  radicals_error_retry: { vi: "✕ Lỗi, thử lại", en: "✕ Error, retry" },
  radicals_update_default: { vi: "🔄 Cập nhật mặc định", en: "🔄 Update Default" },
  radicals_is_default: { vi: "★ Đang là mặc định", en: "★ Currently Default" },
  radicals_set_default: { vi: "⭐ Đặt làm mặc định", en: "⭐ Set as Default" },
  radicals_stroke_count: (n) => ({ vi: `${n} nét`, en: `${n} stroke${n === "1" ? "" : "s"}` }),
  radicals_stroke_unknown: { vi: "chưa xác định số nét", en: "stroke count unknown" },
  radicals_lists_field_label: { vi: "Danh sách", en: "Lists" },
  stroke_modal_title: { vi: "Thứ tự nét bút", en: "Stroke Order" },
  stroke_modal_replay: { vi: "▶ Xem lại", en: "▶ Replay" },
  stroke_modal_no_data: (char) => ({
    vi: `Chưa có dữ liệu nét bút cho chữ "${char}" trong nguồn dữ liệu.`,
    en: `No stroke data available for "${char}" in the data source.`,
  }),
  components_breakdown_label: { vi: "Bộ thủ cấu thành", en: "Components" },
  view_stroke_order_of: (char) => ({ vi: `Xem thứ tự nét bút của ${char}`, en: `View stroke order of ${char}` }),

  // Hán tự (Character list panel)
  hanzi_panel_title: { vi: "Danh sách Hán tự trong kho dữ liệu", en: "Character List in Storage" },
  hanzi_panel_title_guest: { vi: "Đăng nhập để thêm chữ của riêng bạn", en: "Sign in to add your own characters" },
  hanzi_search_placeholder: {
    vi: "Tìm chữ theo Hán tự, pinyin, nghĩa, hoặc Hán Việt…",
    en: "Search by character, pinyin, meaning, or Sino-Vietnamese…",
  },
  hanzi_all_status: { vi: "Tất cả trạng thái", en: "All statuses" },
  hanzi_is_official: { vi: "★ Đang là mặc định", en: "★ Currently default" },
  hanzi_is_pending: { vi: "⭐ Đặt làm mặc định", en: "⭐ Pending publish" },
  hanzi_export_excel: { vi: "⬇ Xuất Excel", en: "⬇ Export Excel" },
  hanzi_export_none: { vi: "Không có chữ nào để xuất.", en: "No characters to export." },
  hanzi_export_success: (n) => ({ vi: `Đã xuất ${n} chữ ra file Excel.`, en: `Exported ${n} characters to Excel.` }),
  hanzi_export_fail: { vi: "Xuất Excel thất bại. Vui lòng thử lại.", en: "Excel export failed. Please try again." },
  hanzi_count: (filtered, total) => ({ vi: `${filtered} / ${total} chữ`, en: `${filtered} / ${total} characters` }),

  // Từ vựng (Word list panel)
  vocab_panel_title: { vi: "Danh sách từ vựng trong kho dữ liệu", en: "Vocabulary List in Storage" },
  vocab_panel_title_guest: { vi: "Đăng nhập để thêm từ của riêng bạn", en: "Sign in to add your own words" },
  vocab_empty: {
    vi: 'Bạn chưa có từ nào. Hãy thêm từ ở tab "Tạo thẻ từ mới".',
    en: 'You have no words yet. Add some in the "Create New Card" tab.',
  },
  vocab_search_placeholder: {
    vi: "Tìm từ theo Hán tự, pinyin, nghĩa, hoặc Hán Việt…",
    en: "Search by character, pinyin, meaning, or Sino-Vietnamese…",
  },
  vocab_export_none: { vi: "Không có từ nào để xuất.", en: "No words to export." },
  vocab_export_success: (n) => ({ vi: `Đã xuất ${n} từ ra file Excel.`, en: `Exported ${n} words to Excel.` }),
  vocab_count: (filtered, total) => ({ vi: `${filtered} / ${total} từ`, en: `${filtered} / ${total} words` }),
};

function t(key, meaningDisplay, ...args) {
  const entry = UI_TEXT[key];
  if (!entry) return key;
  const resolved = typeof entry === "function" ? entry(...args) : entry;
  const val = meaningDisplay === "en" ? resolved.en ?? resolved.vi : resolved.vi;
  return val;
}

// Radical list names like "1 nét" / "11-17 nét" are stored as literal data
// values (same as any other list name, e.g. "HSK1"), not routed through the
// t() dictionary. This translates just that one predictable pattern for
// display in English mode, without touching the underlying stored value --
// filtering by list still works correctly since the real name is unchanged.
function displayListName(name, meaningDisplay) {
  if (meaningDisplay !== "en") return name;
  if (name === "Chưa phân loại") return "Uncategorized";
  const range = name.match(/^(\d+)-(\d+) nét$/);
  if (range) return `${range[1]}-${range[2]} strokes`;
  const single = name.match(/^(\d+) nét$/);
  if (single) return `${single[1]} stroke${single[1] === "1" ? "" : "s"}`;
  return name;
}

function uid() {
  return Math.random().toString(36).slice(2, 10);
}

/* ---------- Supabase row <-> app object conversion ---------- */
function rowToBushou(row) {
  return { char: row.char, pinyin: row.pinyin, meaning: row.meaning, sv: row.sv, strokes: row.strokes, lists: row.lists || [] };
}
function bushouToRow(b, userId) {
  return {
    user_id: userId,
    char: b.char,
    pinyin: b.pinyin,
    meaning: b.meaning,
    sv: b.sv,
    strokes: typeof b.strokes === "number" ? b.strokes : null,
    lists: b.lists || [],
  };
}
function rowToChar(row) {
  return {
    char: row.char,
    pinyin: row.pinyin,
    meaning: row.meaning,
    meaning_vi: row.meaning_vi || "",
    sv: row.sv,
    components: row.components || [],
    lists: row.lists || [],
  };
}
function charToRow(c, userId) {
  return {
    user_id: userId,
    char: c.char,
    pinyin: c.pinyin,
    meaning: c.meaning,
    meaning_vi: c.meaning_vi || "",
    sv: c.sv,
    components: c.components || [],
    lists: c.lists || [],
  };
}
function rowToWord(row) {
  return {
    word: row.word,
    chars: row.chars || [],
    pinyin: row.pinyin,
    meaning: row.meaning,
    meaning_vi: row.meaning_vi || "",
    sv: row.sv,
    lists: row.lists || [],
  };
}
function wordToRow(w, userId) {
  return {
    user_id: userId,
    word: w.word,
    chars: w.chars || [],
    pinyin: w.pinyin,
    meaning: w.meaning,
    meaning_vi: w.meaning_vi || "",
    sv: w.sv,
    lists: w.lists || [],
  };
}
function charToOfficialRow(c) {
  return {
    char: c.char,
    pinyin: c.pinyin,
    meaning: c.meaning,
    meaning_vi: c.meaning_vi || "",
    sv: c.sv,
    components: c.components || [],
    lists: c.lists || [],
  };
}
function wordToOfficialRow(w) {
  return {
    word: w.word,
    chars: w.chars || [],
    pinyin: w.pinyin,
    meaning: w.meaning,
    meaning_vi: w.meaning_vi || "",
    sv: w.sv,
    lists: w.lists || [],
  };
}

function shuffle(arr) {
  const a = [...arr];
  for (let i = a.length - 1; i > 0; i--) {
    const j = Math.floor(Math.random() * (i + 1));
    [a[i], a[j]] = [a[j], a[i]];
  }
  return a;
}

// A character can belong to more than one list. Supports the current `lists`
// array field, and falls back to the older single `list` string for any
// data saved before multi-list support existed.
function getLists(c) {
  if (Array.isArray(c.lists) && c.lists.length > 0) return c.lists;
  if (c.list && typeof c.list === "string") return [c.list];
  return ["Chưa phân loại"];
}

// A deck bundles together (content_type, list_name) pairs rather than
// storing items directly -- this resolves that bundle into the actual pool
// of practiceable items at study time, tagged with their type so progress
// tracking (which already namespaces by "char:x" / "word:x" / "bushou:x")
// stays correctly separated even when a deck mixes all three types.
function resolveDeckItems(deck, { characterList, wordList, bushouList }) {
  if (!deck) return [];
  const items = [];
  const seen = new Set(); // avoid duplicates if multiple bundled lists overlap
  (deck.lists || []).forEach(({ content_type, list_name }) => {
    if (content_type === "char") {
      characterList.forEach((c) => {
        if (getLists(c).includes(list_name) && !seen.has(`char:${c.char}`)) {
          seen.add(`char:${c.char}`);
          items.push({ type: "char", key: c.char, data: c });
        }
      });
    } else if (content_type === "word") {
      wordList.forEach((w) => {
        if ((w.lists || []).includes(list_name) && !seen.has(`word:${w.word}`)) {
          seen.add(`word:${w.word}`);
          items.push({ type: "word", key: w.word, data: w });
        }
      });
    } else if (content_type === "bushou") {
      (bushouList || []).forEach((b) => {
        if ((b.lists || []).includes(list_name) && !seen.has(`bushou:${b.char}`)) {
          seen.add(`bushou:${b.char}`);
          items.push({ type: "bushou", key: b.char, data: b });
        }
      });
    }
  });
  return items;
}

/* ---------- Mizige (米字格) target grid — the signature element ---------- */
function CharacterGrid({ children, state, size = 168 }) {
  const borderColor =
    state === "correct" ? COLORS.bamboo : state === "wrong" ? COLORS.error : state === "revealed" ? COLORS.gold : COLORS.grid;
  const mid = size / 2;
  const inset = size * 0.024; // matches the original 4px inset at 168px
  const far = size - inset;
  return (
    <div
      style={{
        width: size,
        height: size,
        position: "relative",
        border: `2.5px solid ${borderColor}`,
        background: COLORS.card,
        borderRadius: 6,
        boxShadow: "inset 0 0 0 1px rgba(43,41,37,0.04)",
        transition: "border-color 0.25s ease",
        flexShrink: 0,
      }}
    >
      <svg
        width={size}
        height={size}
        style={{ position: "absolute", inset: 0, opacity: 0.55 }}
      >
        <line x1={mid} y1={inset} x2={mid} y2={far} stroke={COLORS.grid} strokeWidth="1" strokeDasharray="4 4" />
        <line x1={inset} y1={mid} x2={far} y2={mid} stroke={COLORS.grid} strokeWidth="1" strokeDasharray="4 4" />
        <line x1={inset} y1={inset} x2={far} y2={far} stroke={COLORS.grid} strokeWidth="1" strokeDasharray="3 5" />
        <line x1={far} y1={inset} x2={inset} y2={far} stroke={COLORS.grid} strokeWidth="1" strokeDasharray="3 5" />
      </svg>
      <div
        style={{
          position: "absolute",
          inset: 0,
          display: "flex",
          alignItems: "center",
          justifyContent: "center",
          flexWrap: "wrap",
          gap: 2,
          padding: 10,
        }}
      >
        {children}
      </div>
    </div>
  );
}

/* ---------- A single bushou chip (palette or selection) ---------- */
/* ---------- Shared two-box meaning display (English + Vietnamese),
   respecting the viewer's meaningDisplay preference. Used everywhere a
   character/word's meaning is shown. ---------- */
/* ---------- Same EN/VI preference as MeaningBoxes, but as a plain string
   for compact inline sentences (status messages) rather than boxes. ---------- */
function formatMeaningInline(meaning, meaningVi, meaningDisplay) {
  const en = meaning || "";
  const vi = meaningVi || "";
  if (meaningDisplay === "en") return en;
  if (meaningDisplay === "vi") return vi || en;
  if (en && vi) return `${en} / ${vi}`;
  return en || vi;
}


function MeaningBoxes({ meaning, meaningVi, meaningDisplay, large }) {
  const showEn = meaningDisplay !== "vi";
  const showVi = meaningDisplay !== "en";
  const valueSize = large ? 15 : 11.5;
  const labelSize = large ? 10 : 9;
  const boxStyle = {
    padding: large ? "5px 10px" : "3px 7px",
    borderRadius: 6,
    background: COLORS.chipBg,
    border: `1px solid ${COLORS.hairline}`,
    textAlign: "center",
  };
  return (
    <div style={{ display: "flex", gap: 6, justifyContent: "center", flexWrap: "wrap" }}>
      {showEn && (
        <div style={boxStyle}>
          <div style={{ fontSize: labelSize, color: COLORS.inkSoft, fontWeight: 700, textTransform: "uppercase", letterSpacing: 0.3 }}>EN</div>
          <div style={{ fontSize: valueSize, color: COLORS.ink }}>{meaning || "—"}</div>
        </div>
      )}
      {showVi && (
        <div style={boxStyle}>
          <div style={{ fontSize: labelSize, color: COLORS.inkSoft, fontWeight: 700, textTransform: "uppercase", letterSpacing: 0.3 }}>VI</div>
          <div style={{ fontSize: valueSize, color: meaningVi ? COLORS.ink : COLORS.inkSoft, fontStyle: meaningVi ? "normal" : "italic" }}>
            {meaningVi || "(chưa dịch)"}
          </div>
        </div>
      )}
    </div>
  );
}


function Chip({ info, onClick, disabled, big, tone, meaningDisplay }) {
  const tooltip = info
    ? [info.pinyin, meaningDisplay !== "vi" ? info.meaning : null, meaningDisplay !== "en" ? `SV: ${info.sv}` : null]
        .filter(Boolean)
        .join(" · ")
    : "";
  return (
    <button
      onClick={onClick}
      disabled={disabled}
      title={tooltip}
      className="hanzi-chip"
      style={{
        fontFamily: "'Noto Serif SC', 'STKaiti', 'Kaiti SC', serif",
        fontSize: big ? 30 : 26,
        width: big ? 56 : 48,
        height: big ? 56 : 48,
        lineHeight: 1,
        display: "flex",
        alignItems: "center",
        justifyContent: "center",
        borderRadius: 11,
        border: `1.5px solid ${tone === "seal" ? COLORS.seal : COLORS.gold}`,
        background: disabled ? "rgba(80,133,165,0.12)" : COLORS.chipBg,
        color: disabled ? COLORS.inkSoft : COLORS.ink,
        cursor: disabled ? "default" : "pointer",
        opacity: disabled ? 0.35 : 1,
        transition: "transform 0.12s ease, box-shadow 0.12s ease",
        boxShadow: disabled ? "none" : "0 1px 0 rgba(43,41,37,0.08)",
      }}
    >
      {info ? info.char : "?"}
    </button>
  );
}

// Every lookup call (auto-fill) needs a real login — this fetches the
// current Supabase access token to send as Authorization, or returns null
// if there's no session (guest), so callers know to prompt for login
// instead of even attempting the request.
async function getAuthHeaders() {
  const { data } = await supabase.auth.getSession();
  const token = data && data.session ? data.session.access_token : null;
  if (!token) return null;
  return { Authorization: `Bearer ${token}` };
}

/* ---------- Shown when a guest tries to use a lookup (auto-fill) —
   lookups cost real money per call, so they require a real account. ---------- */
function LanguagePromptModal({ onChoose }) {
  return (
    <div
      style={{
        position: "fixed",
        inset: 0,
        background: "rgba(30,28,10,0.55)",
        zIndex: 1400,
        display: "flex",
        alignItems: "center",
        justifyContent: "center",
        padding: 20,
      }}
    >
      <div
        style={{
          background: COLORS.card,
          borderRadius: 14,
          padding: "28px 24px",
          width: "90%",
          maxWidth: 380,
          textAlign: "center",
          boxShadow: "0 12px 40px rgba(0,0,0,0.35)",
        }}
      >
        <div style={{ fontFamily: "'Noto Serif SC', serif", fontSize: 28, fontWeight: 700, color: COLORS.ink, marginBottom: 12 }}>
          学部首学汉字
        </div>
        <div style={{ fontSize: 14.5, color: COLORS.ink, marginBottom: 4, fontWeight: 600 }}>
          Chọn ngôn ngữ hiển thị
        </div>
        <div style={{ fontSize: 14.5, color: COLORS.ink, marginBottom: 20, fontWeight: 600 }}>
          Choose your display language
        </div>
        <div style={{ display: "flex", flexDirection: "column", gap: 10 }}>
          <button
            type="button"
            onClick={() => onChoose("both")}
            className="seal-btn"
            style={{ ...sealBtnStyle, padding: "12px 20px", fontSize: 14 }}
          >
            Tiếng Việt + English
          </button>
          <button
            type="button"
            onClick={() => onChoose("vi")}
            className="ghost-btn"
            style={{ ...ghostBtnStyle, padding: "11px 20px", fontSize: 14 }}
          >
            Tiếng Việt
          </button>
          <button
            type="button"
            onClick={() => onChoose("en")}
            className="ghost-btn"
            style={{ ...ghostBtnStyle, padding: "11px 20px", fontSize: 14 }}
          >
            English
          </button>
        </div>
      </div>
    </div>
  );
}

function AuthRequiredModal({ onClose, onSignIn }) {
  return (
    <div
      onClick={onClose}
      style={{
        position: "fixed",
        inset: 0,
        background: "rgba(30,28,10,0.55)",
        zIndex: 1300,
        display: "flex",
        alignItems: "center",
        justifyContent: "center",
        padding: 20,
      }}
    >
      <div
        onClick={(e) => e.stopPropagation()}
        style={{
          background: COLORS.card,
          borderRadius: 14,
          padding: "26px 24px",
          width: "90%",
          maxWidth: 340,
          textAlign: "center",
          boxShadow: "0 12px 40px rgba(0,0,0,0.35)",
        }}
      >
        <div style={{ fontSize: 17, fontWeight: 700, color: COLORS.ink, marginBottom: 8 }}>Cần đăng nhập</div>
        <div style={{ fontSize: 13, color: COLORS.inkSoft, marginBottom: 18, lineHeight: 1.5 }}>
          Tính năng tra cứu tự động (🔍) cần tài khoản đã đăng nhập. Vui lòng đăng nhập hoặc đăng ký để sử dụng.
        </div>
        <div style={{ display: "flex", justifyContent: "center", gap: 8 }}>
          <button
            type="button"
            onClick={() => {
              onClose();
              onSignIn && onSignIn();
            }}
            className="seal-btn"
            style={{ ...sealBtnStyle, padding: "8px 18px", fontSize: 13 }}
          >
            Đăng nhập / Đăng ký
          </button>
          <button type="button" onClick={onClose} className="ghost-btn" style={{ ...ghostBtnStyle, padding: "8px 18px", fontSize: 13 }}>
            Đóng
          </button>
        </div>
      </div>
    </div>
  );
}

/* ---------- Shown when a logged-in user hits their lookup quota. ---------- */
function LimitReachedModal({ onClose, count, limit, tier, reason, onViewPremium }) {
  const isDisabled = reason === "DISABLED";
  return (
    <div
      onClick={onClose}
      style={{
        position: "fixed",
        inset: 0,
        background: "rgba(30,28,10,0.55)",
        zIndex: 1300,
        display: "flex",
        alignItems: "center",
        justifyContent: "center",
        padding: 20,
      }}
    >
      <div
        onClick={(e) => e.stopPropagation()}
        style={{
          background: COLORS.card,
          borderRadius: 14,
          padding: "26px 24px",
          width: "90%",
          maxWidth: 340,
          textAlign: "center",
          boxShadow: "0 12px 40px rgba(0,0,0,0.35)",
        }}
      >
        <div style={{ fontSize: 17, fontWeight: 700, color: COLORS.error, marginBottom: 8 }}>
          {isDisabled ? "Tài khoản đã bị vô hiệu hóa" : "Đã hết lượt tra cứu"}
        </div>
        {isDisabled ? (
          <div style={{ fontSize: 13, color: COLORS.inkSoft, marginBottom: 6, lineHeight: 1.5 }}>
            Tài khoản của bạn hiện không thể sử dụng tính năng tra cứu tự động. Vui lòng liên hệ quản trị viên nếu bạn cho rằng đây là nhầm lẫn.
          </div>
        ) : (
          <div style={{ fontSize: 13, color: COLORS.inkSoft, marginBottom: 6, lineHeight: 1.5 }}>
            Bạn đã dùng {count}/{limit} lượt tra cứu tự động ở gói {tier}.
          </div>
        )}
        <div style={{ fontSize: 12.5, color: COLORS.inkSoft, marginBottom: 18, lineHeight: 1.5 }}>
          Bạn vẫn có thể thêm chữ/từ thủ công (không cần tra cứu tự động) mà không bị giới hạn.
          {!isDisabled && " Nâng cấp để có thêm lượt tra cứu."}
        </div>
        <div style={{ display: "flex", justifyContent: "center", gap: 8 }}>
          {!isDisabled && onViewPremium && (
            <button
              type="button"
              onClick={() => {
                onClose();
                onViewPremium();
              }}
              className="seal-btn"
              style={{ ...sealBtnStyle, padding: "8px 18px", fontSize: 13 }}
            >
              Xem gói nâng cấp
            </button>
          )}
          <button type="button" onClick={onClose} className="ghost-btn" style={{ ...ghostBtnStyle, padding: "8px 18px", fontSize: 13 }}>
            Đã hiểu
          </button>
        </div>
      </div>
    </div>
  );
}

/* ---------- Shown when a user tries to browse/play a list their tier or
   course doesn't unlock. ---------- */
function ListLockedModal({ onClose, listName, onViewPremium }) {
  return (
    <div
      onClick={onClose}
      style={{
        position: "fixed",
        inset: 0,
        background: "rgba(30,28,10,0.55)",
        zIndex: 1300,
        display: "flex",
        alignItems: "center",
        justifyContent: "center",
        padding: 20,
      }}
    >
      <div
        onClick={(e) => e.stopPropagation()}
        style={{
          background: COLORS.card,
          borderRadius: 14,
          padding: "26px 24px",
          width: "90%",
          maxWidth: 340,
          textAlign: "center",
          boxShadow: "0 12px 40px rgba(0,0,0,0.35)",
        }}
      >
        <div style={{ fontSize: 17, fontWeight: 700, color: COLORS.gold, marginBottom: 8 }}>Danh sách yêu cầu nâng cấp</div>
        <div style={{ fontSize: 13, color: COLORS.inkSoft, marginBottom: 18, lineHeight: 1.5 }}>
          Danh sách "{listName}" chỉ dành cho một số gói thành viên hoặc khóa học nhất định. Nâng cấp tài khoản hoặc liên hệ quản trị viên để được cấp quyền truy cập.
        </div>
        <div style={{ display: "flex", justifyContent: "center", gap: 8 }}>
          <button
            type="button"
            onClick={() => {
              onClose();
              onViewPremium && onViewPremium();
            }}
            className="seal-btn"
            style={{ ...sealBtnStyle, padding: "8px 18px", fontSize: 13 }}
          >
            Xem gói nâng cấp
          </button>
          <button type="button" onClick={onClose} className="ghost-btn" style={{ ...ghostBtnStyle, padding: "8px 18px", fontSize: 13 }}>
            Đóng
          </button>
        </div>
      </div>
    </div>
  );
}

/* ---------- Stroke order animation, shared by both radicals and characters.
   Uses HanziWriter + its default Make Me a Hanzi data source (a real,
   dictionary-derived stroke database fetched on demand) — NOT anything
   generated here, since stroke order needs to be actually correct. Some
   rare/side-form radicals (e.g. 忄, 扌) may not exist in that dataset;
   that's handled as a clean "no data available" message rather than a
   guess. ---------- */
function StrokeOrderModal({ char, onClose, meaningDisplay }) {
  const targetRef = useRef(null);
  const writerRef = useRef(null);
  const [status, setStatus] = useState("loading"); // loading | ready | error

  useEffect(() => {
    let cancelled = false;
    if (!targetRef.current) return;
    setStatus("loading");

    try {
      const writer = HanziWriter.create(targetRef.current, char, {
        width: 260,
        height: 260,
        padding: 12,
        showOutline: true,
        strokeAnimationSpeed: 1,
        delayBetweenStrokes: 300,
        strokeColor: COLORS.ink,
        outlineColor: COLORS.grid,
        radicalColor: COLORS.seal,
        onLoadCharDataSuccess: () => {
          if (cancelled) return;
          setStatus("ready");
          writer.animateCharacter();
        },
        onLoadCharDataError: () => {
          if (!cancelled) setStatus("error");
        },
      });
      writerRef.current = writer;
    } catch (e) {
      console.error("Stroke order failed to load:", e);
      setStatus("error");
    }

    return () => {
      cancelled = true;
    };
    // eslint-disable-next-line react-hooks/exhaustive-deps
  }, [char]);

  function replay() {
    if (writerRef.current) writerRef.current.animateCharacter();
  }

  return (
    <div
      onClick={onClose}
      style={{
        position: "fixed",
        inset: 0,
        background: "rgba(30,28,10,0.55)",
        zIndex: 1100,
        display: "flex",
        alignItems: "center",
        justifyContent: "center",
        padding: 20,
      }}
    >
      <div
        onClick={(e) => e.stopPropagation()}
        style={{
          background: COLORS.card,
          borderRadius: 14,
          padding: "24px 22px",
          width: "90%",
          maxWidth: 340,
          textAlign: "center",
          position: "relative",
          boxShadow: "0 12px 40px rgba(0,0,0,0.35)",
        }}
      >
        <button
          type="button"
          onClick={onClose}
          title="Đóng"
          style={{
            position: "absolute",
            top: 10,
            right: 10,
            width: 26,
            height: 26,
            lineHeight: "24px",
            fontSize: 14,
            border: `1px solid ${COLORS.hairline}`,
            borderRadius: "50%",
            background: COLORS.chipBg,
            color: COLORS.inkSoft,
            cursor: "pointer",
          }}
        >
          ✕
        </button>

        <div style={{ fontSize: 11, color: COLORS.inkSoft, textTransform: "uppercase", letterSpacing: 0.6, marginBottom: 10 }}>
          {t("stroke_modal_title", meaningDisplay)}
        </div>

        <div
          ref={targetRef}
          style={{
            width: 260,
            height: 260,
            margin: "0 auto",
            border: `1px solid ${COLORS.hairline}`,
            borderRadius: 6,
            background: COLORS.card,
            display: status === "error" ? "none" : "block",
          }}
        />

        {status === "error" && (
          <div style={{ width: 260, height: 260, margin: "0 auto", display: "flex", alignItems: "center", justifyContent: "center", color: COLORS.inkSoft, fontSize: 13, padding: 16 }}>
            {t("stroke_modal_no_data", meaningDisplay, char)}
          </div>
        )}

        <button
          type="button"
          onClick={replay}
          disabled={status === "error"}
          className="ghost-btn"
          style={{ ...ghostBtnStyle, marginTop: 14, opacity: status === "error" ? 0.4 : 1 }}
        >
          {t("stroke_modal_replay", meaningDisplay)}
        </button>
      </div>
    </div>
  );
}


class ErrorBoundary extends React.Component {
  constructor(props) {
    super(props);
    this.state = { error: null };
  }
  static getDerivedStateFromError(error) {
    return { error };
  }
  componentDidCatch(error, info) {
    console.error("HanziBuilder crashed:", error, info);
  }
  render() {
    if (this.state.error) {
      return (
        <div
          style={{
            fontFamily: "'Noto Sans', 'Inter', sans-serif",
            padding: 30,
            maxWidth: 600,
            margin: "40px auto",
            background: "#FFFFFF",
            border: "2px solid #C7C9C5",
            borderRadius: 14,
            color: "#16181A",
          }}
        >
          <div style={{ fontWeight: 700, marginBottom: 8, color: "#A8482F" }}>Đã xảy ra lỗi · Something went wrong</div>
          <div style={{ fontSize: 13, marginBottom: 14, color: "#33383B" }}>
            {String(this.state.error && this.state.error.message ? this.state.error.message : this.state.error)}
          </div>
          <button
            onClick={() => this.setState({ error: null })}
            style={{
              background: "#A8482F",
              border: "none",
              color: "#FBF9EF",
              borderRadius: 7,
              padding: "9px 18px",
              fontSize: 13.5,
              cursor: "pointer",
            }}
          >
            Thử lại · Retry
          </button>
        </div>
      );
    }
    return this.props.children;
  }
}

export default function HanziBuilder({ userId, userEmail, onRequireAuth }) {
  return (
    <ErrorBoundary>
      <HanziBuilderApp userId={userId} userEmail={userEmail} onRequireAuth={onRequireAuth} />
    </ErrorBoundary>
  );
}

function HanziBuilderApp({ userId, userEmail, onRequireAuth }) {
  const [customBushou, setCustomBushou] = useState([]);
  const [customChars, setCustomChars] = useState([]);
  const [customWords, setCustomWords] = useState([]);
  const [deletedChars, setDeletedChars] = useState([]);
  const [needsReview, setNeedsReview] = useState([]);
  const [loaded, setLoaded] = useState(false);
  const [tab, setTab] = useState("home");
  // Lets other tabs (Combine Radicals, Flashcards, Handwriting) jump
  // straight into Library's "Add New Cards" sub-tab, not just Library
  // generally. LibraryTab reads this once and clears it, so a normal
  // visit to Library afterward doesn't keep jumping back to Add.
  const [pendingLibrarySubTab, setPendingLibrarySubTab] = useState(null);
  function goToAddNewCards() {
    setTab("library");
    setPendingLibrarySubTab("add");
  }

  // The shared default data, loaded from Supabase for EVERYONE (including
  // guests, via public SELECT policies) so admin corrections go live for
  // every visitor without a code deploy. Falls back to the hardcoded
  // SEED_ arrays if the tables are empty/unreachable, so the app never
  // breaks even if this fetch has a problem.
  const [officialBushou, setOfficialBushou] = useState(null); // null = not loaded yet
  const [officialChars, setOfficialChars] = useState(null);
  const [officialWords, setOfficialWords] = useState(null);
  const [decks, setDecks] = useState([]); // [{ id, name, description, lists: [{content_type, list_name}] }]
  const [isAdmin, setIsAdmin] = useState(false);
  const [lookupCount, setLookupCount] = useState(0);
  const [lookupLimit, setLookupLimit] = useState(100);
  const [tier, setTier] = useState("Free");
  const [courseName, setCourseName] = useState(null);
  const [meaningDisplay, setMeaningDisplay] = useState("both"); // 'both' | 'en' | 'vi'

  const updateMeaningDisplay = useCallback(
    async (next) => {
      setMeaningDisplay(next);
      if (!userId) return; // guest: session-only, nothing to persist
      const { error } = await supabase.from("profiles").update({ meaning_display: next }).eq("user_id", userId);
      if (error) console.error("Could not save meaning display preference:", error);
    },
    [userId]
  );

  // Ask first-time guest visitors to pick a language up front, rather than
  // leaving them to find the small toggle on their own. Only for guests --
  // logged-in users already have a saved preference (or a sensible default)
  // tied to their account. Remembered in localStorage so it only asks once
  // per browser, not on every visit.
  const [showLangPrompt, setShowLangPrompt] = useState(false);
  useEffect(() => {
    if (userId) return; // logged in -- never show this to them
    let alreadyChosen = false;
    try {
      alreadyChosen = localStorage.getItem("hanzi_guest_lang_chosen") === "1";
    } catch (e) {
      // localStorage unavailable (e.g. private browsing) -- just skip the prompt
      alreadyChosen = true;
    }
    if (!alreadyChosen) setShowLangPrompt(true);
  }, [userId]);

  function handleLangPromptChoice(choice) {
    updateMeaningDisplay(choice);
    setShowLangPrompt(false);
    try {
      localStorage.setItem("hanzi_guest_lang_chosen", "1");
    } catch (e) {
      // ignore -- worst case it asks again next visit
    }
  }

  // List access rules -- loaded for everyone, including guests, since list
  // NAMES are meant to be visible to everyone (that's the upgrade hook).
  // null = not loaded yet.
  const [listSettings, setListSettings] = useState(null); // [{name, admin_only, allowed_tiers}]
  const [listCourseAccess, setListCourseAccess] = useState(null); // [{list_name, course_name}]

  useEffect(() => {
    let cancelled = false;
    (async () => {
      const [lsRes, lcaRes] = await Promise.all([
        supabase.from("list_settings").select("*"),
        supabase.from("list_course_access").select("*"),
      ]);
      if (cancelled) return;
      setListSettings(!lsRes.error ? lsRes.data || [] : []);
      setListCourseAccess(!lcaRes.error ? lcaRes.data || [] : []);
    })();
    return () => {
      cancelled = true;
    };
  }, []);

  useEffect(() => {
    let cancelled = false;
    (async () => {
      try {
        const [bRes, cRes, wRes] = await Promise.all([
          supabase.from("official_bushou").select("*"),
          supabase.from("official_characters").select("*"),
          supabase.from("official_words").select("*"),
        ]);
        if (cancelled) return;
        setOfficialBushou(!bRes.error && bRes.data.length > 0 ? bRes.data.map(rowToBushou) : SEED_BUSHOU);
        setOfficialChars(!cRes.error && cRes.data.length > 0 ? cRes.data.map(rowToChar) : SEED_CHARACTERS);
        setOfficialWords(!wRes.error && wRes.data.length > 0 ? wRes.data.map(rowToWord) : SEED_WORDS);
      } catch (e) {
        console.error("Could not load official data, using built-in defaults:", e);
        if (!cancelled) {
          setOfficialBushou(SEED_BUSHOU);
          setOfficialChars(SEED_CHARACTERS);
          setOfficialWords(SEED_WORDS);
        }
      }
    })();
    return () => {
      cancelled = true;
    };
  }, []);

  const loadDecks = useCallback(async () => {
    const [decksRes, listsRes] = await Promise.all([
      supabase.from("decks").select("*").order("name"),
      supabase.from("deck_lists").select("*"),
    ]);
    if (decksRes.error || listsRes.error) {
      console.error("Could not load decks:", decksRes.error || listsRes.error);
      return;
    }
    const byDeck = new Map();
    (listsRes.data || []).forEach((row) => {
      if (!byDeck.has(row.deck_id)) byDeck.set(row.deck_id, []);
      byDeck.get(row.deck_id).push({ content_type: row.content_type, list_name: row.list_name });
    });
    setDecks(
      (decksRes.data || []).map((d) => ({
        id: d.id,
        name: d.name,
        description: d.description || "",
        userId: d.user_id || null, // null = shared/official deck; otherwise a personal deck
        lists: byDeck.get(d.id) || [],
      }))
    );
  }, []);

  useEffect(() => {
    loadDecks();
    // Re-fetch on login/logout so personal decks appear/disappear correctly
    // (RLS already scopes the query, this just re-runs it).
    // eslint-disable-next-line react-hooks/exhaustive-deps
  }, [loadDecks, userId]);

  // Unread/new counts for notification badges -- light polling rather
  // than true realtime, so badges update on their own every ~45 seconds
  // without needing a page reload, without the added complexity of a
  // persistent live connection. adminBadges tracks each admin section
  // separately so it's visible at a glance *where* new activity is, not
  // just that some exists somewhere.
  const [unreadForUser, setUnreadForUser] = useState(0);
  const [adminBadges, setAdminBadges] = useState({ messages: 0, feedback: 0, suggestions: 0, comments: 0 });
  const unreadForAdmin = adminBadges.messages + adminBadges.feedback + adminBadges.suggestions + adminBadges.comments;

  useEffect(() => {
    if (!userId) {
      setUnreadForUser(0);
      setAdminBadges({ messages: 0, feedback: 0, suggestions: 0, comments: 0 });
      return;
    }
    let cancelled = false;
    async function checkUnread() {
      const { count: userCount } = await supabase
        .from("messages")
        .select("id", { count: "exact", head: true })
        .eq("user_id", userId)
        .eq("sender", "admin")
        .eq("read", false);
      if (!cancelled) setUnreadForUser(userCount || 0);

      if (isAdmin) {
        const [messagesRes, feedbackRes, suggestionsRes, commentsRes] = await Promise.all([
          supabase.from("messages").select("id", { count: "exact", head: true }).eq("sender", "user").eq("read", false),
          supabase.from("feedback").select("id", { count: "exact", head: true }).eq("read", false),
          supabase.from("card_suggestions").select("id", { count: "exact", head: true }).eq("status", "new"),
          supabase.from("blog_comments").select("id", { count: "exact", head: true }).eq("status", "pending"),
        ]);
        if (!cancelled) {
          setAdminBadges({
            messages: messagesRes.count || 0,
            feedback: feedbackRes.count || 0,
            suggestions: suggestionsRes.count || 0,
            comments: commentsRes.count || 0,
          });
        }
      }
    }
    checkUnread();
    const interval = setInterval(checkUnread, 45000);
    return () => {
      cancelled = true;
      clearInterval(interval);
    };
  }, [userId, isAdmin]);

  useEffect(() => {
    if (!userId) {
      setIsAdmin(false);
      return;
    }
    let cancelled = false;
    (async () => {
      // Ensure a profiles row exists (harmless no-op if it already does),
      // and keep the email fresh for the admin panel to search by. is_admin
      // itself is never set through the app.
      await supabase.from("profiles").upsert({ user_id: userId, email: userEmail || null }, { onConflict: "user_id" });
      const { data, error } = await supabase
        .from("profiles")
        .select("is_admin, lookup_count, lookup_limit, tier, course_name, meaning_display")
        .eq("user_id", userId)
        .single();
      if (cancelled) return;
      setIsAdmin(!error && data ? !!data.is_admin : false);
      if (!error && data) {
        setLookupCount(data.lookup_count || 0);
        setLookupLimit(data.lookup_limit != null ? data.lookup_limit : 1000);
        setTier(data.tier || "Free");
        setCourseName(data.course_name || null);
        setMeaningDisplay(data.meaning_display || "both");
      }
    })();
    return () => {
      cancelled = true;
    };
  }, [userId, userEmail]);

  /* ---------- load this user's data from Supabase (skipped for guests) ---------- */
  useEffect(() => {
    if (!userId) {
      // Guest mode: no account, nothing to load — play with seed data only,
      // any additions stay in memory for this browser tab and are not saved.
      setCustomBushou([]);
      setCustomChars([]);
      setDeletedChars([]);
      setNeedsReview([]);
      setLoaded(true);
      return;
    }

    let cancelled = false;

    (async () => {
      const [bushouRes, charsRes, wordsRes, deletedRes, reviewRes] = await Promise.all([
        supabase.from("custom_bushou").select("*").eq("user_id", userId),
        supabase.from("custom_characters").select("*").eq("user_id", userId),
        supabase.from("custom_words").select("*").eq("user_id", userId),
        supabase.from("deleted_characters").select("char").eq("user_id", userId),
        supabase.from("needs_review").select("char").eq("user_id", userId),
      ]);

      if (cancelled) return;

      if (bushouRes.error) console.error("Load bushou failed:", bushouRes.error);
      else setCustomBushou(bushouRes.data.map(rowToBushou));

      if (charsRes.error) console.error("Load characters failed:", charsRes.error);
      else setCustomChars(charsRes.data.map(rowToChar));

      if (wordsRes.error) console.error("Load words failed:", wordsRes.error);
      else setCustomWords(wordsRes.data.map(rowToWord));

      if (deletedRes.error) console.error("Load deleted list failed:", deletedRes.error);
      else setDeletedChars(deletedRes.data.map((r) => r.char));

      if (reviewRes.error) console.error("Load review list failed:", reviewRes.error);
      else setNeedsReview(reviewRes.data.map((r) => r.char));

      setLoaded(true);
    })();

    return () => {
      cancelled = true;
    };
  }, [userId]);

  // Add (or overwrite) a single radical for this user. Local state and the
  // database are updated together so the UI feels instant.
  const addBushouRow = useCallback(
    async (entry) => {
      setCustomBushou((prev) => {
        const without = prev.filter((b) => b.char !== entry.char);
        return [...without, entry];
      });
      if (!userId) return; // guest mode: keep in memory only, nothing to save
      const { error } = await supabase
        .from("custom_bushou")
        .upsert(bushouToRow(entry, userId), { onConflict: "user_id,char" });
      if (error) console.error("Could not save radical:", error);
    },
    [userId]
  );

  const addCharacterRow = useCallback(
    async (entry) => {
      setCustomChars((prev) => [...prev, entry]);
      // If this char was previously hidden (deleted from this user's own
      // view), adding it back is a clear signal to un-hide it too —
      // otherwise the old tombstone would keep masking it forever.
      setDeletedChars((prev) => prev.filter((ch) => ch !== entry.char));
      if (!userId) return; // guest mode: keep in memory only, nothing to save
      const { error } = await supabase
        .from("custom_characters")
        .upsert(charToRow(entry, userId), { onConflict: "user_id,char" });
      if (error) console.error("Could not save character:", error);
      const { error: delError } = await supabase
        .from("deleted_characters")
        .delete()
        .eq("user_id", userId)
        .eq("char", entry.char);
      if (delError) console.error("Could not clear hidden flag:", delError);
    },
    [userId]
  );

  const updateCharacterRow = useCallback(
    async (char, updatedFields) => {
      setCustomChars((prev) => {
        const existing = prev.find((c) => c.char === char);
        const base = existing || characterList.find((c) => c.char === char) || { char };
        const merged = { ...base, ...updatedFields, char };
        return existing ? prev.map((c) => (c.char === char ? merged : c)) : [...prev, merged];
      });
      if (!userId) return; // guest mode: keep in memory only, nothing to save
      // Read the merged record back out of state on the next tick isn't safe (stale closure),
      // so recompute the same merge here for what we send to the database.
      const existing = customChars.find((c) => c.char === char);
      const base = existing || characterList.find((c) => c.char === char) || { char };
      const merged = { ...base, ...updatedFields, char };
      const { error } = await supabase
        .from("custom_characters")
        .upsert(charToRow(merged, userId), { onConflict: "user_id,char" });
      if (error) console.error("Could not update character:", error);
    },
    // eslint-disable-next-line react-hooks/exhaustive-deps
    [userId, customChars]
  );

  const deleteCharacterRow = useCallback(
    async (char) => {
      setCustomChars((prev) => prev.filter((c) => c.char !== char));
      setDeletedChars((prev) => (prev.includes(char) ? prev : [...prev, char]));
      if (!userId) return; // guest mode: keep in memory only, nothing to save
      await supabase.from("custom_characters").delete().eq("user_id", userId).eq("char", char);
      const { error } = await supabase
        .from("deleted_characters")
        .upsert({ user_id: userId, char }, { onConflict: "user_id,char" });
      if (error) console.error("Could not record deletion:", error);
    },
    [userId]
  );

  const addWordRow = useCallback(
    async (entry) => {
      setCustomWords((prev) => {
        const without = prev.filter((w) => w.word !== entry.word);
        return [...without, entry];
      });
      if (!userId) return; // guest mode: keep in memory only, nothing to save
      const { error } = await supabase
        .from("custom_words")
        .upsert(wordToRow(entry, userId), { onConflict: "user_id,word" });
      if (error) console.error("Could not save word:", error);
    },
    [userId]
  );

  const deleteWordRow = useCallback(
    async (word) => {
      setCustomWords((prev) => prev.filter((w) => w.word !== word));
      if (!userId) return; // guest mode: keep in memory only, nothing to save
      const { error } = await supabase.from("custom_words").delete().eq("user_id", userId).eq("word", word);
      if (error) console.error("Could not delete word:", error);
    },
    [userId]
  );

  // "Promote to default": push the current (possibly personally-corrected)
  // version of an item into the shared official_* table, so every visitor
  // sees it — then clean up the now-redundant personal copy, since the
  // official entry already reflects it.
  const promoteBushouToDefault = useCallback(
    async (b) => {
      const { error } = await supabase
        .from("official_bushou")
        .upsert(
          { char: b.char, pinyin: b.pinyin, meaning: b.meaning, sv: b.sv, strokes: typeof b.strokes === "number" ? b.strokes : null },
          { onConflict: "char" }
        );
      if (error) {
        console.error("Could not promote radical to default:", error);
        throw error;
      }
      setOfficialBushou((prev) => {
        const without = (prev || []).filter((x) => x.char !== b.char);
        return [...without, b];
      });
      setCustomBushou((prev) => prev.filter((x) => x.char !== b.char));
      if (userId) await supabase.from("custom_bushou").delete().eq("user_id", userId).eq("char", b.char);
    },
    [userId]
  );

  // Reverse of promote: pulls an entry OUT of the shared default table and
  // hands it back to the admin personally, so it doesn't just vanish.
  const withdrawBushouFromDefault = useCallback(
    async (b) => {
      const { error } = await supabase.from("official_bushou").delete().eq("char", b.char);
      if (error) {
        console.error("Could not withdraw radical from default:", error);
        throw error;
      }
      setOfficialBushou((prev) => (prev || []).filter((x) => x.char !== b.char));
      setCustomBushou((prev) => {
        const without = prev.filter((x) => x.char !== b.char);
        return [...without, b];
      });
      if (userId) {
        await supabase
          .from("custom_bushou")
          .upsert({ user_id: userId, char: b.char, pinyin: b.pinyin, meaning: b.meaning, sv: b.sv, strokes: typeof b.strokes === "number" ? b.strokes : null }, { onConflict: "user_id,char" });
      }
    },
    [userId]
  );

  const promoteCharacterToDefault = useCallback(
    async (c) => {
      const { error } = await supabase.from("official_characters").upsert(charToOfficialRow(c), { onConflict: "char" });
      if (error) {
        console.error("Could not promote character to default:", error);
        throw error;
      }
      setOfficialChars((prev) => {
        const without = (prev || []).filter((x) => x.char !== c.char);
        return [...without, c];
      });
      setCustomChars((prev) => prev.filter((x) => x.char !== c.char));
      if (userId) await supabase.from("custom_characters").delete().eq("user_id", userId).eq("char", c.char);
    },
    [userId]
  );

  const withdrawCharacterFromDefault = useCallback(
    async (c) => {
      const { error } = await supabase.from("official_characters").delete().eq("char", c.char);
      if (error) {
        console.error("Could not withdraw character from default:", error);
        throw error;
      }
      setOfficialChars((prev) => (prev || []).filter((x) => x.char !== c.char));
      setCustomChars((prev) => {
        const without = prev.filter((x) => x.char !== c.char);
        return [...without, c];
      });
      if (userId) await supabase.from("custom_characters").upsert(charToRow(c, userId), { onConflict: "user_id,char" });
    },
    [userId]
  );

  // Permanent removal from the shared default — unlike withdraw, this does
  // NOT hand a personal copy back. Used when an admin's "delete" should
  // actually mean "gone for everyone," not just "gone from my own view."
  const deleteCharacterFromOfficial = useCallback(async (char) => {
    const { error } = await supabase.from("official_characters").delete().eq("char", char);
    if (error) {
      console.error("Could not delete character from default:", error);
      throw error;
    }
    setOfficialChars((prev) => (prev || []).filter((x) => x.char !== char));
  }, []);

  const promoteWordToDefault = useCallback(
    async (w) => {
      const { error } = await supabase.from("official_words").upsert(wordToOfficialRow(w), { onConflict: "word" });
      if (error) {
        console.error("Could not promote word to default:", error);
        throw error;
      }
      setOfficialWords((prev) => {
        const without = (prev || []).filter((x) => x.word !== w.word);
        return [...without, w];
      });
      setCustomWords((prev) => prev.filter((x) => x.word !== w.word));
      if (userId) await supabase.from("custom_words").delete().eq("user_id", userId).eq("word", w.word);
    },
    [userId]
  );

  const withdrawWordFromDefault = useCallback(
    async (w) => {
      const { error } = await supabase.from("official_words").delete().eq("word", w.word);
      if (error) {
        console.error("Could not withdraw word from default:", error);
        throw error;
      }
      setOfficialWords((prev) => (prev || []).filter((x) => x.word !== w.word));
      setCustomWords((prev) => {
        const without = prev.filter((x) => x.word !== w.word);
        return [...without, w];
      });
      if (userId) await supabase.from("custom_words").upsert(wordToRow(w, userId), { onConflict: "user_id,word" });
    },
    [userId]
  );

  const deleteWordFromOfficial = useCallback(async (word) => {
    const { error } = await supabase.from("official_words").delete().eq("word", word);
    if (error) {
      console.error("Could not delete word from default:", error);
      throw error;
    }
    setOfficialWords((prev) => (prev || []).filter((x) => x.word !== word));
  }, []);

  const persistNeedsReview = useCallback(
    async (next, char, adding) => {
      setNeedsReview(next);
      if (!userId) return; // guest mode: keep in memory only, nothing to save
      if (adding) {
        const { error } = await supabase
          .from("needs_review")
          .upsert({ user_id: userId, char }, { onConflict: "user_id,char" });
        if (error) console.error("Could not add to review list:", error);
      } else {
        const { error } = await supabase
          .from("needs_review")
          .delete()
          .eq("user_id", userId)
          .eq("char", char);
        if (error) console.error("Could not clear review flag:", error);
      }
    },
    [userId]
  );

  const bushouList = useMemo(() => {
    const map = new Map();
    [...(officialBushou || SEED_BUSHOU), ...customBushou].forEach((b) => map.set(b.char, b));
    return Array.from(map.values());
  }, [officialBushou, customBushou]);

  // List names that are completely private to admin -- draft/testing
  // content, invisible to everyone else (not just gated by tier).
  const adminOnlyListNames = useMemo(() => {
    return new Set((listSettings || []).filter((s) => s.admin_only).map((s) => s.name));
  }, [listSettings]);

  const characterList = useMemo(() => {
    const map = new Map();
    [...(officialChars || SEED_CHARACTERS), ...customChars].forEach((c) => map.set(c.char, c));
    deletedChars.forEach((ch) => map.delete(ch));
    let arr = Array.from(map.values());
    if (!isAdmin && adminOnlyListNames.size > 0) {
      arr = arr.filter((c) => {
        const lists = getLists(c);
        return lists.length === 0 || lists.some((l) => !adminOnlyListNames.has(l));
      });
    }
    return arr;
  }, [officialChars, customChars, deletedChars, isAdmin, adminOnlyListNames]);

  const wordList = useMemo(() => {
    const map = new Map();
    [...(officialWords || SEED_WORDS), ...customWords].forEach((w) => map.set(w.word, w));
    let arr = Array.from(map.values());
    if (!isAdmin && adminOnlyListNames.size > 0) {
      arr = arr.filter((w) => {
        const lists = w.lists || [];
        return lists.length === 0 || lists.some((l) => !adminOnlyListNames.has(l));
      });
    }
    return arr;
  }, [officialWords, customWords, isAdmin, adminOnlyListNames]);

  // Whether the CURRENT user can see a given list's actual contents
  // (separate from admin_only, which hides it entirely and is handled
  // above) -- gated by tier, or by course for Enrolled Course students.
  const checkListAccess = useCallback(
    (listName) => {
      if (isAdmin) return true;
      const setting = (listSettings || []).find((s) => s.name === listName);
      if (!setting) return true; // not configured -> open to everyone
      if (setting.admin_only) return false;
      if (setting.allowed_tiers && setting.allowed_tiers.includes(tier)) return true;
      if (tier === "Enrolled Course" && courseName) {
        const hasCourseGrant = (listCourseAccess || []).some(
          (g) => g.list_name === listName && g.course_name === courseName
        );
        if (hasCourseGrant) return true;
      }
      return false;
    },
    [isAdmin, tier, courseName, listSettings, listCourseAccess]
  );

  // Every list name currently in use anywhere (characters or words) --
  // note this runs on the ADMIN's own characterList/wordList, which is
  // unfiltered (the admin_only filter only applies to non-admins), so this
  // correctly includes admin-only lists too, for managing them.
  const allListNamesInUse = useMemo(() => {
    const set = new Set();
    characterList.forEach((c) => getLists(c).forEach((l) => set.add(l.trim())));
    wordList.forEach((w) => (w.lists || []).forEach((l) => set.add(l.trim())));
    return Array.from(set).sort((a, b) => a.localeCompare(b, "vi"));
  }, [characterList, wordList]);

  // Which chars/words are currently live for everyone (vs. only in this
  // admin's personal data) — drives the promote/withdraw toggle state.
  const officialBushouKeys = useMemo(() => new Set((officialBushou || []).map((b) => b.char)), [officialBushou]);
  const officialCharKeys = useMemo(() => new Set((officialChars || []).map((c) => c.char)), [officialChars]);
  const officialWordKeys = useMemo(() => new Set((officialWords || []).map((w) => w.word)), [officialWords]);

  // Whether THIS admin has a personal edit sitting on top of (or instead
  // of) the official value — i.e. there's something new to publish, as
  // opposed to just viewing the already-published default with no changes.
  const overrideBushouKeys = useMemo(() => new Set(customBushou.map((b) => b.char)), [customBushou]);
  const overrideCharKeys = useMemo(() => new Set(customChars.map((c) => c.char)), [customChars]);
  const overrideWordKeys = useMemo(() => new Set(customWords.map((w) => w.word)), [customWords]);

  const findBushou = useCallback(
    (ch) => bushouList.find((b) => b.char === ch) || { char: ch, pinyin: "—", meaning: "unknown", sv: "—" },
    [bushouList]
  );

  return (
    <div
      style={{
        fontFamily: "'Noto Sans', 'Inter', sans-serif",
        background: COLORS.paper,
        minHeight: "100%",
        color: COLORS.ink,
        padding: "28px 16px 48px",
      }}
    >
      <style>{`
        ${FONT_IMPORT}
        .hanzi-chip:not(:disabled):hover { transform: translateY(-2px); box-shadow: 0 4px 10px rgba(43,41,37,0.15); }
        .hanzi-chip:not(:disabled):active { transform: translateY(0); }
        .tab-btn { transition: color 0.15s ease, border-color 0.15s ease; }
        .seal-btn { transition: filter 0.15s ease, transform 0.1s ease; }
        .seal-btn:hover { filter: brightness(1.08); }
        .seal-btn:active { transform: translateY(1px); }
        .ghost-btn:hover { background: rgba(43,41,37,0.06) !important; }
        @keyframes popIn { 0% { transform: scale(0.7); opacity: 0; } 100% { transform: scale(1); opacity: 1; } }
        .pop { animation: popIn 0.28s cubic-bezier(.2,1.4,.4,1) both; }
        input, select, textarea { font-family: 'Noto Sans', 'Inter', sans-serif; }
        .rich-text-editor:empty:before { content: attr(data-placeholder); color: ${COLORS.metadata}; }
        .rich-text-editor a { color: ${COLORS.seal}; }
        .cjk-enhanced { font-family: 'KaiTi', 'STKaiti', 'Kaiti SC', 'Noto Serif SC', serif; font-size: 1.2em; }
        .home-feature-card { cursor: pointer; transition: transform 0.18s ease, box-shadow 0.18s ease; }
        .home-feature-card:hover { transform: scale(1.05); box-shadow: 0 6px 18px rgba(0,0,0,0.15); }
        ::selection { background: ${COLORS.gold}55; }
        @media (max-width: 480px) {
          .field-row { flex-direction: column; align-items: flex-start !important; gap: 4px !important; }
          .field-row-label { width: auto !important; }
          .field-row input, .field-row select { width: 100% !important; box-sizing: border-box; }
          .list-pills-row { padding-left: 0 !important; }
          .autofill-hint { padding-left: 0 !important; }
          .side-nav-layout { flex-direction: column !important; }
          .side-nav-menu { flex-direction: row !important; overflow-x: auto; width: 100% !important; gap: 6px !important; }
          .side-nav-menu button { white-space: nowrap; }
        }
      `}</style>

      {showLangPrompt && <LanguagePromptModal onChoose={handleLangPromptChoice} />}

      <div style={{ position: "fixed", top: 44, right: 16, zIndex: 40 }}>
        <MeaningDisplayToggle value={meaningDisplay} onChange={updateMeaningDisplay} />
      </div>

      <div style={{ maxWidth: 760, margin: "0 auto" }}>
        <Header meaningDisplay={meaningDisplay} />
        {userId && <LookupQuotaBadge count={lookupCount} limit={lookupLimit} tier={tier} isAdmin={isAdmin} meaningDisplay={meaningDisplay} />}
        <Tabs tab={tab} setTab={setTab} isAdmin={isAdmin} meaningDisplay={meaningDisplay} unreadForUser={unreadForUser} unreadForAdmin={unreadForAdmin} />

        {!loaded ? (
          <div style={{ textAlign: "center", padding: 60, color: COLORS.inkSoft }}>{t("loading", meaningDisplay)}</div>
        ) : tab === "home" ? (
          <HomeTab
            setTab={setTab}
            userId={userId}
            tier={tier}
            lookupCount={lookupCount}
            lookupLimit={lookupLimit}
            onRequireAuth={onRequireAuth}
            meaningDisplay={meaningDisplay}
          />
        ) : tab === "play" ? (
          <PlayTab
            characterList={characterList}
            wordList={wordList}
            bushouList={bushouList}
            findBushou={findBushou}
            needsReview={needsReview}
            onMarkNeedsReview={(char) => {
              if (!needsReview.includes(char)) persistNeedsReview([...needsReview, char], char, true);
            }}
            onClearNeedsReview={(char) => {
              if (needsReview.includes(char))
                persistNeedsReview(needsReview.filter((c) => c !== char), char, false);
            }}
            isAdmin={isAdmin}
            checkListAccess={checkListAccess}
            onViewPremium={() => setTab("premium")}
            meaningDisplay={meaningDisplay}
            goToAddNewCards={goToAddNewCards}
          />
        ) : tab === "flashcards" ? (
          <FlashcardsTab
            userId={userId}
            characterList={characterList}
            wordList={wordList}
            bushouList={bushouList}
            decks={decks}
            isAdmin={isAdmin}
            checkListAccess={checkListAccess}
            onRequireAuth={onRequireAuth}
            onViewPremium={() => setTab("premium")}
            meaningDisplay={meaningDisplay}
            goToAddNewCards={goToAddNewCards}
          />
        ) : tab === "writing" ? (
          <WritingPracticeTab
            characterList={characterList}
            bushouList={bushouList}
            decks={decks}
            isAdmin={isAdmin}
            checkListAccess={checkListAccess}
            onViewPremium={() => setTab("premium")}
            meaningDisplay={meaningDisplay}
            goToAddNewCards={goToAddNewCards}
          />
        ) : tab === "library" ? (
          <LibraryTab
            bushouList={bushouList}
            characterList={characterList}
            wordList={wordList}
            customWords={customWords}
            onAddCharacter={addCharacterRow}
            onAddBushou={addBushouRow}
            officialBushouKeys={officialBushouKeys}
            overrideBushouKeys={overrideBushouKeys}
            onPromoteBushou={promoteBushouToDefault}
            onWithdrawBushou={withdrawBushouFromDefault}
            onDeleteCharacter={deleteCharacterRow}
            onDeleteCharacterFromOfficial={deleteCharacterFromOfficial}
            onUpdateCharacter={updateCharacterRow}
            officialCharKeys={officialCharKeys}
            overrideCharKeys={overrideCharKeys}
            onPromoteCharacter={promoteCharacterToDefault}
            onWithdrawCharacter={withdrawCharacterFromDefault}
            findBushou={findBushou}
            onAddWord={addWordRow}
            onDeleteWord={deleteWordRow}
            onDeleteWordFromOfficial={deleteWordFromOfficial}
            officialWordKeys={officialWordKeys}
            overrideWordKeys={overrideWordKeys}
            onPromoteWord={promoteWordToDefault}
            onWithdrawWord={withdrawWordFromDefault}
            isAdmin={isAdmin}
            checkListAccess={checkListAccess}
            onViewPremium={() => setTab("premium")}
            meaningDisplay={meaningDisplay}
            userId={userId}
            onRequireAuth={onRequireAuth}
            onQuotaUpdate={(count, limit) => {
              setLookupCount(count);
              if (typeof limit === "number") setLookupLimit(limit);
            }}
            pendingSubTab={pendingLibrarySubTab}
            onConsumePendingSubTab={() => setPendingLibrarySubTab(null)}
          />
        ) : tab === "management" ? (
          <ManagementTab
            userId={userId}
            isAdmin={isAdmin}
            tier={tier}
            lookupCount={lookupCount}
            lookupLimit={lookupLimit}
            courseName={courseName}
            characterList={characterList}
            wordList={wordList}
            bushouList={bushouList}
            decks={decks}
            onDecksChanged={loadDecks}
            meaningDisplay={meaningDisplay}
            onRequireAuth={onRequireAuth}
          />
        ) : tab === "premium" ? (
          <PremiumTab meaningDisplay={meaningDisplay} />
        ) : tab === "blog" ? (
          <BlogTab meaningDisplay={meaningDisplay} />
        ) : tab === "about" ? (
          <AboutTab meaningDisplay={meaningDisplay} />
        ) : tab === "privacy" ? (
          <PrivacyTab />
        ) : tab === "terms" ? (
          <TermsTab />
        ) : tab === "feedback" ? (
          <FeedbackTab meaningDisplay={meaningDisplay} userId={userId} />
        ) : tab === "admin" ? (
          <AdminPanel
            isAdmin={isAdmin}
            allListNamesInUse={allListNamesInUse}
            meaningDisplay={meaningDisplay}
            characterList={characterList}
            wordList={wordList}
            bushouList={bushouList}
            decks={decks}
            onDecksChanged={loadDecks}
            userId={userId}
            adminBadges={adminBadges}
          />
        ) : null}
      </div>
      <SiteFooter setTab={setTab} meaningDisplay={meaningDisplay} />
    </div>
  );
}

/* ---------- Header ---------- */
function Header({ meaningDisplay }) {
  return (
    <div style={{ display: "flex", alignItems: "center", justifyContent: "center", gap: "clamp(10px, 3vw, 20px)", marginBottom: 22, flexWrap: "nowrap" }}>
      <img
        src={LOGO_DATA_URI}
        alt="minouQ - Scholastic Joy"
        style={{ width: "clamp(90px, 26vw, 260px)", height: "auto", flexShrink: 0 }}
      />
      <div style={{ textAlign: "left", minWidth: 0 }}>
        <div
          style={{
            fontFamily: "'Noto Serif SC', 'STKaiti', 'Kaiti SC', serif",
            fontSize: "clamp(18px, 4.5vw, 34px)",
            fontWeight: 700,
            letterSpacing: 3,
            color: COLORS.ink,
          }}
        >
          学部首学汉字
        </div>
        {meaningDisplay !== "en" && (
          <div
            style={{
              fontFamily: "'Noto Sans', Inter, 'Segoe UI', sans-serif",
              fontSize: "clamp(9px, 1.7vw, 13px)",
              fontWeight: 600,
              color: COLORS.sealDark,
              marginTop: 5,
              maxWidth: 380,
              lineHeight: 1.4,
            }}
          >
            {UI_TEXT.header_subtitle.vi}
          </div>
        )}
        {meaningDisplay !== "vi" && (
          <div
            style={{
              fontFamily: meaningDisplay === "en" ? "'Noto Sans', Inter, 'Segoe UI', sans-serif" : undefined,
              fontSize: meaningDisplay === "en" ? "clamp(9px, 1.7vw, 13px)" : "clamp(8px, 1.3vw, 10px)",
              fontWeight: meaningDisplay === "en" ? 600 : 400,
              color: meaningDisplay === "en" ? COLORS.sealDark : COLORS.inkSoft,
              marginTop: 5,
              letterSpacing: meaningDisplay === "en" ? undefined : 0.2,
              fontStyle: meaningDisplay === "en" ? "normal" : "italic",
              maxWidth: meaningDisplay === "en" ? 380 : 360,
              lineHeight: meaningDisplay === "en" ? 1.4 : 1.5,
            }}
          >
            {UI_TEXT.header_subtitle.en}
          </div>
        )}
      </div>
    </div>
  );
}

/* ---------- Countdown showing how many auto-fill lookups this user has
   left. Admins are unlimited, so it shows their usage without a limit. ---------- */
function LookupQuotaBadge({ count, limit, tier, isAdmin, meaningDisplay }) {
  const remaining = Math.max(0, limit - count);
  const isLow = !isAdmin && remaining <= Math.max(5, limit * 0.1);
  const isOut = !isAdmin && remaining === 0;
  const accentColor = isOut ? COLORS.error : isLow ? COLORS.gold : COLORS.seal;
  const bg = isOut ? "rgba(168,72,47,0.08)" : isLow ? "rgba(80,133,165,0.08)" : "rgba(49,112,142,0.07)";
  return (
    <div style={{ display: "flex", justifyContent: "center", marginBottom: 20 }}>
      <div
        style={{
          display: "inline-flex",
          alignItems: "center",
          gap: 8,
          padding: "7px 15px",
          borderRadius: 999,
          border: `1.5px solid ${accentColor}`,
          background: bg,
        }}
      >
        <span style={{ fontSize: 12, fontWeight: 700, color: accentColor, letterSpacing: 0.3, textTransform: "uppercase" }}>
          {isAdmin ? "Admin" : tier}
        </span>
        <span style={{ width: 1, height: 14, background: accentColor, opacity: 0.35 }} />
        <span style={{ fontSize: 10, fontWeight: 600, color: COLORS.ink }}>
          {isAdmin ? t("quota_admin_usage", meaningDisplay, count) : t("quota_remaining", meaningDisplay, remaining, limit)}
        </span>
      </div>
    </div>
  );
}

/* ---------- Which meaning box(es) to show -- English, Vietnamese, or
   both. Persisted per-account; guests get a session-only choice since
   there's no account to save it to. ---------- */
function MeaningDisplayToggle({ value, onChange }) {
  const options = [
    { id: "both", label: "EN + VI" },
    { id: "en", label: "EN" },
    { id: "vi", label: "VI" },
  ];
  return (
    <div
      style={{
        display: "inline-flex",
        borderRadius: 999,
        border: `1px solid ${COLORS.hairline}`,
        overflow: "hidden",
        background: COLORS.card,
        boxShadow: "0 2px 8px rgba(0,0,0,0.08)",
      }}
    >
      {options.map((opt) => (
        <button
          key={opt.id}
          type="button"
          onClick={() => onChange(opt.id)}
          style={{
            padding: "5px 14px",
            fontSize: 11.5,
            fontWeight: 700,
            border: "none",
            cursor: "pointer",
            background: value === opt.id ? COLORS.seal : "transparent",
            color: value === opt.id ? "#FBF9EF" : COLORS.inkSoft,
          }}
        >
          {opt.label}
        </button>
      ))}
    </div>
  );
}


function Tabs({ tab, setTab, isAdmin, meaningDisplay, unreadForUser, unreadForAdmin }) {
  const items = [
    { id: "home", label: t("tab_home", meaningDisplay) },
    { id: "play", label: t("tab_play", meaningDisplay) },
    { id: "flashcards", label: t("tab_flashcards", meaningDisplay) },
    { id: "writing", label: t("tab_writing", meaningDisplay) },
    { id: "library", label: t("tab_library", meaningDisplay) },
    { id: "management", label: t("tab_management", meaningDisplay), badge: unreadForUser },
  ];
  if (isAdmin) items.push({ id: "admin", label: t("tab_admin", meaningDisplay), badge: unreadForAdmin });
  return (
    <div
      style={{
        display: "flex",
        justifyContent: "center",
        flexWrap: "wrap",
        gap: "4px 26px",
        marginBottom: 26,
        borderBottom: `1px solid ${COLORS.hairline}`,
      }}
    >
      {items.map((it) => (
        <button
          key={it.id}
          className="tab-btn"
          onClick={() => setTab(it.id)}
          style={{
            background: "none",
            border: "none",
            borderBottom: `2px solid ${tab === it.id ? COLORS.seal : "transparent"}`,
            color: tab === it.id ? COLORS.ink : COLORS.inkSoft,
            fontWeight: tab === it.id ? 700 : 600,
            fontSize: 14,
            padding: "10px 2px",
            marginBottom: -1,
            cursor: "pointer",
            position: "relative",
          }}
        >
          {it.label}
          {it.badge > 0 && (
            <span
              style={{
                position: "absolute",
                top: 3,
                right: -14,
                minWidth: 15,
                height: 15,
                borderRadius: 999,
                background: COLORS.error,
                color: "#FBF9EF",
                fontSize: 9.5,
                fontWeight: 700,
                display: "flex",
                alignItems: "center",
                justifyContent: "center",
                padding: "0 3px",
              }}
            >
              {it.badge > 9 ? "9+" : it.badge}
            </span>
          )}
        </button>
      ))}
    </div>
  );
}

/* ================= PLAY TAB ================= */
const REVIEW_LIST_VALUE = "__needs_review__";

function HomeTab({ setTab, userId, tier, lookupCount, lookupLimit, onRequireAuth, meaningDisplay }) {
  return (
    <div style={{ textAlign: "center" }}>
      <div style={{ fontSize: 30, fontWeight: 800, color: COLORS.ink, lineHeight: 1.25, marginBottom: 12 }}>
        {t("home_headline", meaningDisplay)}
      </div>
      <div style={{ fontSize: 14.5, color: COLORS.inkSoft, maxWidth: 420, margin: "0 auto 22px", lineHeight: 1.6, whiteSpace: "pre-line" }}>
        {t("home_subtext", meaningDisplay)}
      </div>

      {!userId ? (
        <button
          type="button"
          onClick={() => onRequireAuth && onRequireAuth()}
          style={{
            background: COLORS.seal,
            color: "#FBF9EF",
            border: "none",
            borderRadius: 999,
            padding: "11px 30px",
            fontSize: 14.5,
            fontWeight: 700,
            cursor: "pointer",
            marginBottom: 32,
          }}
        >
          {t("home_signup_button", meaningDisplay)}
        </button>
      ) : (
        <div style={{ marginBottom: 32 }}>
          <span
            style={{
              display: "inline-block",
              border: `1px solid ${COLORS.grid}`,
              borderRadius: 999,
              padding: "7px 16px",
              fontSize: 12.5,
              color: COLORS.inkSoft,
            }}
          >
            <span style={{ fontWeight: 700, color: COLORS.seal }}>{tier || "Free"}</span>
            {" · "}
            {t("home_status_lookups", meaningDisplay, lookupCount ?? 0, lookupLimit ?? 1000)}
          </span>
        </div>
      )}

      <div style={{ display: "grid", gridTemplateColumns: "repeat(auto-fit, minmax(150px, 1fr))", gap: 14 }}>
        <CombineRadicalsPreviewCard setTab={setTab} meaningDisplay={meaningDisplay} />
        <FlashcardsPreviewCard setTab={setTab} meaningDisplay={meaningDisplay} />
        <HandwritingPreviewCard setTab={setTab} meaningDisplay={meaningDisplay} />
      </div>
    </div>
  );
}

function CombineRadicalsPreviewCard({ setTab, meaningDisplay }) {
  const targetRef = useRef(null);
  const aRef = useRef(null);
  const bRef = useRef(null);
  const decoyRefs = useRef([]);

  useEffect(() => {
    const CYCLE = 8200;
    function run() {
      const target = targetRef.current, a = aRef.current, b = bRef.current;
      const decoys = decoyRefs.current;
      if (!target || !a || !b) return;
      target.style.transition = "none";
      target.style.opacity = "0.2";
      // Resting state: no offset at all -- flexbox alone decides each
      // character's position, so this can never drift off-center again.
      a.style.transition = "none"; a.style.transform = "translate(0px, 0px)"; a.style.opacity = "1";
      b.style.transition = "none"; b.style.transform = "translate(0px, 0px)"; b.style.opacity = "1";
      decoys.forEach((d) => d && (d.style.opacity = "1"));

      const t1 = setTimeout(() => {
        a.style.transition = "transform 0.7s ease"; b.style.transition = "transform 0.7s ease";
        // Relative nudges toward the target above, regardless of exactly
        // where flexbox placed each character -- always converges near
        // center since the row itself is centered.
        a.style.transform = "translate(-62px, -54px) scale(0.9)";
        b.style.transform = "translate(40px, -54px) scale(0.9)";
        decoys.forEach((d) => { if (d) { d.style.transition = "opacity 0.5s ease"; d.style.opacity = "0.35"; } });
      }, 900);
      const t2 = setTimeout(() => {
        a.style.opacity = "0"; b.style.opacity = "0";
        target.style.transition = "opacity 0.4s ease"; target.style.opacity = "1";
      }, 1900);
      const t3 = setTimeout(() => decoys.forEach((d) => d && (d.style.opacity = "1")), 5200);
      return [t1, t2, t3];
    }
    let timeouts = run();
    const interval = setInterval(() => {
      timeouts.forEach(clearTimeout);
      timeouts = run();
    }, CYCLE);
    return () => {
      clearInterval(interval);
      timeouts.forEach(clearTimeout);
    };
  }, []);

  const charStyle = { fontFamily: "'Noto Serif SC', 'STKaiti', 'Kaiti SC', serif", fontSize: 18, color: "#E1F0F5" };

  return (
    <div
      className="home-feature-card"
      onClick={() => setTab("play")}
      style={{ background: COLORS.seal, borderRadius: 14, padding: 18, height: 220, display: "flex", flexDirection: "column", justifyContent: "space-between", overflow: "hidden" }}
    >
      <div style={{ fontSize: 14, fontWeight: 700, color: "#FBF9EF" }}>{t("tab_play", meaningDisplay)}</div>
      <div style={{ position: "relative", height: 110 }}>
        <div ref={targetRef} style={{ fontFamily: "'Noto Serif SC', 'STKaiti', 'Kaiti SC', serif", fontSize: 38, color: "#FBF9EF", opacity: 0.2, textAlign: "center", transition: "opacity 0.4s ease" }}>好</div>
        <div style={{ display: "flex", justifyContent: "center", alignItems: "center", gap: 4, marginTop: 12 }}>
          <div ref={(el) => (decoyRefs.current[0] = el)} style={charStyle}>木</div>
          <div ref={bRef} style={{ ...charStyle, position: "relative" }}>子</div>
          <div ref={(el) => (decoyRefs.current[1] = el)} style={charStyle}>口</div>
          <div ref={(el) => (decoyRefs.current[2] = el)} style={charStyle}>水</div>
          <div ref={(el) => (decoyRefs.current[3] = el)} style={charStyle}>手</div>
          <div ref={(el) => (decoyRefs.current[4] = el)} style={charStyle}>心</div>
          <div ref={aRef} style={{ ...charStyle, position: "relative" }}>女</div>
        </div>
      </div>
      <div style={{ fontSize: 11, color: "#E1F0F5" }}>{t("home_card_radicals_desc", meaningDisplay)}</div>
    </div>
  );
}

function FlashcardsPreviewCard({ setTab, meaningDisplay }) {
  const cardRef = useRef(null);

  useEffect(() => {
    const CYCLE = 8200;
    function run() {
      const c = cardRef.current;
      if (!c) return;
      c.style.transform = "rotateY(180deg)";
      const timeout = setTimeout(() => { c.style.transform = "rotateY(0deg)"; }, 4500);
      return timeout;
    }
    let timeout = run();
    const interval = setInterval(() => {
      clearTimeout(timeout);
      timeout = run();
    }, CYCLE);
    return () => {
      clearInterval(interval);
      clearTimeout(timeout);
    };
  }, []);

  return (
    <div
      className="home-feature-card"
      onClick={() => setTab("flashcards")}
      style={{ background: "#54697A", borderRadius: 14, padding: 18, height: 220, display: "flex", flexDirection: "column", justifyContent: "space-between" }}
    >
      <div style={{ fontSize: 14, fontWeight: 700, color: "#FBF9EF" }}>{t("tab_flashcards", meaningDisplay)}</div>
      <div style={{ display: "flex", justifyContent: "center", perspective: 500 }}>
        <div ref={cardRef} style={{ width: 108, height: 82, position: "relative", transformStyle: "preserve-3d", transition: "transform 0.7s ease" }}>
          <div style={{ position: "absolute", inset: 0, background: "#FBF9EF", borderRadius: 8, display: "flex", alignItems: "center", justifyContent: "center", backfaceVisibility: "hidden", fontSize: 16, color: "#3A4A56", fontWeight: 700 }}>hǎo</div>
          <div style={{ position: "absolute", inset: 0, background: "#FBF9EF", borderRadius: 8, display: "flex", alignItems: "center", justifyContent: "center", backfaceVisibility: "hidden", transform: "rotateY(180deg)", fontFamily: "'Noto Serif SC', 'STKaiti', 'Kaiti SC', serif", fontSize: 38, color: "#3A4A56" }}>好</div>
        </div>
      </div>
      <div style={{ fontSize: 11, color: "#E3E9EC" }}>{t("home_card_flashcards_desc", meaningDisplay)}</div>
    </div>
  );
}

function HandwritingPreviewCard({ setTab, meaningDisplay }) {
  const labelRef = useRef(null);
  const dotsRef = useRef(null);
  const svgRef = useRef(null);
  const glyphRef = useRef(null);
  const strokeRefs = useRef({});
  const dashLen = { 1: "34.5", 2: "24.5", 3: "33.5", 4: "40.5" };

  function setRef(prefix, n, el) {
    strokeRefs.current[`${prefix}-${n}`] = el;
  }
  function getRef(prefix, n) {
    return strokeRefs.current[`${prefix}-${n}`];
  }

  useEffect(() => {
    const CYCLE = 8200;

    function resetAll() {
      if (dotsRef.current) dotsRef.current.style.opacity = "0";
      if (svgRef.current) svgRef.current.style.opacity = "1";
      if (glyphRef.current) glyphRef.current.style.opacity = "0";
      ["trace", "recall", "reveal"].forEach((p) => {
        [1, 2, 3, 4].forEach((n) => {
          const el = getRef(p, n);
          if (!el) return;
          el.style.transition = "none";
          el.style.opacity = "0";
          el.style.strokeDashoffset = dashLen[n];
        });
      });
    }

    function drawSet(prefix, startDelay, gap, dur, timeouts) {
      [1, 2, 3, 4].forEach((n, i) => {
        const el = getRef(prefix, n);
        if (!el) return;
        timeouts.push(
          setTimeout(() => {
            el.style.opacity = "1";
            el.style.transition = `stroke-dashoffset ${dur}ms ease`;
            void el.getBoundingClientRect();
            el.style.strokeDashoffset = "0";
          }, startDelay + i * gap)
        );
      });
    }
    function fadeOut(prefix) {
      [1, 2, 3, 4].forEach((n) => {
        const el = getRef(prefix, n);
        if (el) el.style.opacity = "0";
      });
    }

    function run() {
      resetAll();
      const timeouts = [];
      if (labelRef.current) labelRef.current.textContent = t("home_hw_step_trace", meaningDisplay);
      if (dotsRef.current) dotsRef.current.style.opacity = "1";
      drawSet("trace", 300, 500, 420, timeouts);

      timeouts.push(
        setTimeout(() => {
          if (labelRef.current) labelRef.current.textContent = t("home_hw_step_recall", meaningDisplay);
          if (dotsRef.current) dotsRef.current.style.opacity = "0";
          fadeOut("trace");
        }, 2500)
      );
      timeouts.push(setTimeout(() => drawSet("recall", 0, 500, 420, timeouts), 2900));
      timeouts.push(
        setTimeout(() => {
          if (labelRef.current) labelRef.current.textContent = t("home_hw_step_reveal", meaningDisplay);
          fadeOut("recall");
          drawSet("reveal", 0, 280, 250, timeouts);
        }, 5300)
      );
      timeouts.push(
        setTimeout(() => {
          if (svgRef.current) svgRef.current.style.opacity = "0";
          if (glyphRef.current) glyphRef.current.style.opacity = "1";
        }, 6700)
      );
      return timeouts;
    }

    let timeouts = run();
    const interval = setInterval(() => {
      timeouts.forEach(clearTimeout);
      timeouts = run();
    }, CYCLE);
    return () => {
      clearInterval(interval);
      timeouts.forEach(clearTimeout);
    };
    // eslint-disable-next-line react-hooks/exhaustive-deps
  }, [meaningDisplay]);

  const strokePaths = {
    1: "M13 14 Q20 13 30 14 Q40 15 47 13",
    2: "M18 30 Q25 29 30 30 Q35 31 42 29",
    3: "M30 14 Q29 23 31 30 Q31 38 29 47",
    4: "M10 47 Q22 46 30 47 Q40 48 50 46",
  };
  const dots = [
    [13, 14], [47, 13],
    [18, 30], [42, 29],
    [30, 14], [29, 47],
    [10, 47], [50, 46],
  ];

  return (
    <div
      className="home-feature-card"
      onClick={() => setTab("writing")}
      style={{ background: "#445566", borderRadius: 14, padding: 18, height: 220, display: "flex", flexDirection: "column", justifyContent: "space-between", overflow: "hidden" }}
    >
      <div style={{ display: "flex", justifyContent: "space-between", alignItems: "baseline" }}>
        <span style={{ fontSize: 14, fontWeight: 700, color: "#FBF9EF" }}>{t("tab_writing", meaningDisplay)}</span>
        <span ref={labelRef} style={{ fontSize: 10, color: "#DCE3E8" }}>{t("home_hw_step_trace", meaningDisplay)}</span>
      </div>
      <div style={{ display: "flex", justifyContent: "center", position: "relative", height: 90, alignItems: "center" }}>
        <svg ref={svgRef} width="70" height="70" viewBox="0 0 60 60">
          <g ref={dotsRef} opacity="0">
            {dots.map(([cx, cy], i) => (
              <circle key={i} cx={cx} cy={cy} r="2.2" fill="#FBF9EF" />
            ))}
          </g>
          {["trace", "recall", "reveal"].map((prefix) =>
            [1, 2, 3, 4].map((n) => (
              <path
                key={`${prefix}-${n}`}
                ref={(el) => setRef(prefix, n, el)}
                d={strokePaths[n]}
                fill="none"
                stroke={prefix === "trace" ? "#FBF9EF" : prefix === "recall" ? "#C8CDD2" : "#7FB8D8"}
                strokeWidth="3.4"
                strokeLinecap="round"
                strokeDasharray={dashLen[n]}
                opacity="0"
              />
            ))
          )}
        </svg>
        <div ref={glyphRef} style={{ position: "absolute", fontFamily: "'Noto Serif SC', 'STKaiti', 'Kaiti SC', serif", fontSize: 46, color: "#FBF9EF", opacity: 0, transition: "opacity 0.5s ease" }}>王</div>
      </div>
      <div style={{ fontSize: 11, color: "#DCE3E8" }}>{t("home_card_writing_desc", meaningDisplay)}</div>
    </div>
  );
}

function PlayTab({ characterList, wordList, bushouList, findBushou, needsReview, onMarkNeedsReview, onClearNeedsReview, isAdmin, checkListAccess, onViewPremium, meaningDisplay, goToAddNewCards }) {
  const [round, setRound] = useState(null); // { target, palette: [{id,char}] }
  const [selected, setSelected] = useState([]); // array of palette ids, in click order
  const [status, setStatus] = useState("playing"); // playing | correct | wrong | revealed
  const [score, setScore] = useState(0);
  const [streak, setStreak] = useState(0);
  const [selectedList, setSelectedList] = useState("Tất cả");
  const [lockedListName, setLockedListName] = useState(null);
  const [difficulty, setDifficulty] = useState("sieu-kho"); // de | trung-binh | kho | sieu-kho
  const usedRef = useRef(new Set());

  const DIFFICULTY_LEVELS = [
    { id: "de", label: "★", paletteSize: 8 },
    { id: "trung-binh", label: "★★", paletteSize: 12 },
    { id: "kho", label: "★★★", paletteSize: 25 },
    { id: "sieu-kho", label: "★★★★", paletteSize: 36 },
  ];

  // Every playable "thing" - single characters and multi-character words -
  // normalized into one shape: { key, display, pinyin, meaning, sv, lists,
  // charGroups }. charGroups has one entry per character to build, each
  // with its own required components — this is what lets a word render as
  // several boxes instead of one.
  const allCandidates = useMemo(() => {
    const singles = characterList.map((c) => ({
      key: c.char,
      display: c.char,
      pinyin: c.pinyin,
      meaning: c.meaning,
      meaningVi: c.meaning_vi,
      sv: c.sv,
      lists: getLists(c),
      charGroups: buildCharGroups([c.char], characterList),
    }));
    const words = wordList.map((w) => ({
      key: w.word,
      display: w.word,
      pinyin: w.pinyin,
      meaning: w.meaning,
      meaningVi: w.meaning_vi,
      sv: w.sv || "",
      lists: w.lists || ["Cơ bản"],
      charGroups: buildCharGroups(w.chars, characterList),
    }));
    return [...singles, ...words].filter((item) => item.charGroups !== null);
  }, [characterList, wordList]);

  const allLists = useMemo(() => {
    const set = new Set();
    allCandidates.forEach((c) => c.lists.forEach((l) => set.add(l.trim())));
    return Array.from(set).sort((a, b) => a.localeCompare(b, "vi"));
  }, [allCandidates]);

  const playable = useMemo(() => {
    return allCandidates.filter((c) => {
      if (selectedList === REVIEW_LIST_VALUE) return needsReview.includes(c.key);
      if (selectedList === "Tất cả") return true;
      return c.lists.some((l) => l.trim() === selectedList);
    });
  }, [allCandidates, selectedList, needsReview]);

  const buildRound = useCallback(() => {
    if (playable.length === 0) {
      setRound(null);
      return;
    }
    let pool = playable.filter((c) => !usedRef.current.has(c.key));
    if (pool.length === 0) {
      usedRef.current = new Set();
      pool = playable;
    }
    const target = pool[Math.floor(Math.random() * pool.length)];
    usedRef.current.add(target.key);

    const neededComponents = target.charGroups.flatMap((g) => g.components);
    const correctChips = neededComponents.map((ch) => ({ id: uid(), char: ch, correct: true }));

    const PALETTE_SIZE = (DIFFICULTY_LEVELS.find((d) => d.id === difficulty) || DIFFICULTY_LEVELS[3]).paletteSize;
    const neededSet = new Set(neededComponents);
    const distractPool = shuffle(bushouList.filter((b) => !neededSet.has(b.char)));
    const distractCount = Math.min(distractPool.length, Math.max(0, PALETTE_SIZE - neededComponents.length));
    const distractChips = distractPool.slice(0, distractCount).map((b) => ({ id: uid(), char: b.char, correct: false }));

    setRound({ target, palette: shuffle([...correctChips, ...distractChips]) });
    setSelected([]);
    setStatus("playing");
  }, [playable, bushouList, difficulty]);

  useEffect(() => {
    usedRef.current = new Set();
    buildRound();
    // eslint-disable-next-line react-hooks/exhaustive-deps
  }, [selectedList, characterList.length, difficulty]);

  const listPicker = (
    <div style={{ textAlign: "center", marginBottom: 16 }}>
      <div style={{ display: "flex", justifyContent: "center", alignItems: "flex-end", gap: 10, flexWrap: "wrap" }}>
        <select
          value={selectedList}
          onChange={(e) => {
            const next = e.target.value;
            if (!isAdmin && next !== "Tất cả" && next !== REVIEW_LIST_VALUE && checkListAccess && !checkListAccess(next)) {
              setLockedListName(next);
              return;
            }
            setSelectedList(next);
          }}
          style={{ ...selectStyle, width: 220, textAlign: "center", display: "inline-block" }}
        >
          <option value="Tất cả" style={{ background: COLORS.chipBg, color: COLORS.ink, fontWeight: 700 }}>{t("play_all_lists", meaningDisplay)}</option>
          <option value={REVIEW_LIST_VALUE} style={{ background: COLORS.chipBg, color: COLORS.ink, fontWeight: 700 }}>{t("play_review_list", meaningDisplay, needsReview.length)}</option>
          {allLists.map((l) => (
            <option key={l} value={l} style={{ background: COLORS.chipBg, color: COLORS.ink, fontWeight: 700 }}>
              {!isAdmin && checkListAccess && !checkListAccess(l) ? `🔒 ${l}` : displayListName(l, meaningDisplay)}
            </option>
          ))}
        </select>

        <div>
          <div style={{ fontSize: 11, fontWeight: 600, color: COLORS.metadata, textTransform: "uppercase", letterSpacing: 0.5, marginBottom: 4 }}>
            {t("play_difficulty_label", meaningDisplay)}
          </div>
          <select
            value={difficulty}
            onChange={(e) => setDifficulty(e.target.value)}
            style={{ ...selectStyle, width: 140, textAlign: "center", display: "inline-block" }}
          >
            {DIFFICULTY_LEVELS.map((lvl) => (
              <option key={lvl.id} value={lvl.id} style={{ background: COLORS.chipBg, color: COLORS.ink, fontWeight: 700 }}>
                {lvl.label}
              </option>
            ))}
          </select>
        </div>

        <button
          type="button"
          onClick={goToAddNewCards}
          style={{
            border: `1px solid ${COLORS.gold}`,
            borderRadius: 11,
            background: COLORS.gold,
            color: "#FBF9EF",
            fontWeight: 700,
            fontSize: 13,
            padding: "9px 16px",
            cursor: "pointer",
            alignSelf: "flex-end",
          }}
        >
          + {t("tab_add", meaningDisplay)}
        </button>
      </div>
      {lockedListName && (
        <ListLockedModal listName={lockedListName} onClose={() => setLockedListName(null)} onViewPremium={onViewPremium} />
      )}
    </div>
  );

  if (!round) {
    return (
      <div>
        {listPicker}
        <div style={{ textAlign: "center", padding: 50, color: COLORS.inkSoft }}>
          {selectedList === REVIEW_LIST_VALUE
            ? t("play_no_playable_review", meaningDisplay)
            : playable.length === 0
            ? t("play_no_playable_list", meaningDisplay)
            : t("play_no_data", meaningDisplay)}
        </div>
      </div>
    );
  }

  const { target, palette } = round;

  // Cumulative index boundaries so we know which selected chips belong to
  // which box: box i owns selected[boundaries[i] .. boundaries[i+1]).
  const boundaries = [0];
  target.charGroups.forEach((g) => boundaries.push(boundaries[boundaries.length - 1] + g.components.length));
  const totalNeeded = boundaries[boundaries.length - 1];

  function multisetEqual(a, b) {
    if (a.length !== b.length) return false;
    const ca = {}, cb = {};
    a.forEach((x) => (ca[x] = (ca[x] || 0) + 1));
    b.forEach((x) => (cb[x] = (cb[x] || 0) + 1));
    return Object.keys(ca).every((k) => ca[k] === cb[k]) && Object.keys(cb).length === Object.keys(ca).length;
  }

  function selectedCharsForBox(i, fromList) {
    return fromList.slice(boundaries[i], boundaries[i + 1]).map((pid) => palette.find((p) => p.id === pid).char);
  }

  // Per-box correctness, used both for the final overall result and for
  // coloring each box individually (a word can be half-right).
  const boxCorrectness = target.charGroups.map((g, i) => multisetEqual(selectedCharsForBox(i, selected), g.components));

  function handleChipClick(id) {
    if (status !== "playing") return;
    const next = [...selected, id];
    setSelected(next);
    if (next.length === totalNeeded) {
      const allCorrect = target.charGroups.every((g, i) => multisetEqual(selectedCharsForBox(i, next), g.components));
      if (allCorrect) {
        setStatus("correct");
        setScore((s) => s + 10 * target.charGroups.length);
        setStreak((s) => s + 1);
        onClearNeedsReview && onClearNeedsReview(target.key);
      } else {
        setStatus("wrong");
        setStreak(0);
      }
    }
  }

  function handleReset() {
    setSelected([]);
    setStatus("playing");
  }

  function handleShowAnswer() {
    if (status !== "playing") return;
    setSelected([]);
    setStatus("revealed");
    setStreak(0);
    onMarkNeedsReview && onMarkNeedsReview(target.key);
  }

  const answerBreakdown = target.charGroups.map((g) => `${g.components.join(" + ")} = ${g.char}`).join(", ");
  const isWord = target.charGroups.length > 1;

  return (
    <div>
      {listPicker}
      <div style={{ display: "flex", justifyContent: "space-between", alignItems: "center", marginBottom: 18, fontSize: 13, color: COLORS.inkSoft }}>
        <div style={{ display: "flex", gap: 16 }}>
          <span>{t("play_score", meaningDisplay)} <strong style={{ color: COLORS.ink }}>{score}</strong></span>
          <span>{t("play_streak", meaningDisplay)} <strong style={{ color: COLORS.ink }}>{streak}</strong></span>
        </div>
        <span>
          {t("play_playable_count", meaningDisplay, playable.length, isWord || playable.some((p) => p.charGroups.length > 1))}
          {selectedList === REVIEW_LIST_VALUE ? ` ${t("play_needs_review_suffix", meaningDisplay)}` : selectedList !== "Tất cả" ? ` (${selectedList})` : ""}
        </span>
      </div>

      {/* Hint card */}
      <div
        style={{
          background: COLORS.card,
          border: `1px solid ${COLORS.hairline}`,
          borderRadius: 14,
          padding: "16px 20px",
          marginBottom: 22,
          textAlign: "center",
          width: "fit-content",
          maxWidth: "100%",
          margin: "0 auto 22px",
        }}
      >
        <div style={{ marginBottom: 4 }}>
          <MeaningBoxes meaning={target.meaning} meaningVi={target.meaningVi} meaningDisplay={meaningDisplay} large />
        </div>
        <div style={{ display: "flex", justifyContent: "center", gap: 22, marginTop: 8, fontSize: 14.5, flexWrap: "wrap" }}>
          <span style={{ color: COLORS.sealDark }}>{t("play_pinyin", meaningDisplay)} <strong>{target.pinyin}</strong></span>
          {target.sv && meaningDisplay !== "en" && <span style={{ color: COLORS.bamboo }}>{t("play_han_viet", meaningDisplay)} <strong>{target.sv}</strong></span>}
        </div>
      </div>

      {/* Build area: one box per character in the word */}
      <div style={{ display: "flex", justifyContent: "center", gap: 10, marginBottom: 10, flexWrap: "wrap" }}>
        {target.charGroups.map((g, i) => {
          const boxState =
            status === "correct" || status === "revealed"
              ? status
              : status === "wrong"
              ? boxCorrectness[i]
                ? "correct"
                : "wrong"
              : "playing";
          const charsInBox = selectedCharsForBox(i, selected);
          return (
            <CharacterGrid key={i} state={boxState} size={isWord ? 110 : 168}>
              {status === "correct" || status === "revealed" ? (
                <div className="pop" style={{ fontFamily: "'Noto Serif SC', 'STKaiti', 'Kaiti SC', serif", fontSize: isWord ? 54 : 78, color: COLORS.bamboo }}>
                  {g.char}
                </div>
              ) : charsInBox.length === 0 ? (
                i === 0 ? (
                  <div style={{ fontSize: 12, color: COLORS.inkSoft, textAlign: "center" }}>
                    {meaningDisplay === "en" ? (
                      <>select<br />radical below</>
                    ) : (
                      <>chọn<br />bộ thủ bên dưới</>
                    )}
                  </div>
                ) : null
              ) : (
                charsInBox.map((ch, ci) => (
                  <span
                    key={ci}
                    style={{
                      fontFamily: "'Noto Serif SC', 'STKaiti', 'Kaiti SC', serif",
                      fontSize: isWord ? 26 : 40,
                      color: status === "wrong" && !boxCorrectness[i] ? COLORS.error : COLORS.ink,
                    }}
                  >
                    {ch}
                  </span>
                ))
              )}
            </CharacterGrid>
          );
        })}
      </div>

      <div style={{ textAlign: "center", minHeight: 22, marginBottom: 14, fontSize: 13.5, fontWeight: 600 }}>
        {status === "correct" && (
          <span style={{ color: COLORS.gold }}>
            {t("play_correct_prefix", meaningDisplay)} {target.display} ({target.pinyin}) — {formatMeaningInline(target.meaning, target.meaningVi, meaningDisplay)}
          </span>
        )}
        {status === "wrong" && <span style={{ color: COLORS.error }}>{t("play_wrong_answer", meaningDisplay)} {answerBreakdown}</span>}
        {status === "revealed" && (
          <span style={{ color: COLORS.gold }}>
            {t("play_answer_reveal", meaningDisplay)} {answerBreakdown} ({target.pinyin}) — {formatMeaningInline(target.meaning, target.meaningVi, meaningDisplay)}
            {target.sv && meaningDisplay !== "en" ? `, ${t("play_han_viet", meaningDisplay)} ${target.sv}` : ""}
          </span>
        )}
      </div>

      {/* Palette */}
      <div style={{ display: "flex", justifyContent: "center", flexWrap: "wrap", gap: 10, marginBottom: 20 }}>
        {palette.map((p) => (
          <Chip
            key={p.id}
            info={findBushou(p.char)}
            disabled={selected.includes(p.id) || status !== "playing"}
            onClick={() => handleChipClick(p.id)}
          />
        ))}
      </div>

      <div style={{ display: "flex", justifyContent: "center", gap: 10, flexWrap: "wrap" }}>
        {status === "playing" && (
          <button onClick={handleReset} className="ghost-btn" style={{ ...ghostBtnStyle, border: `1.5px solid ${COLORS.ink}`, color: COLORS.ink, fontWeight: 600 }}>
            {t("play_undo", meaningDisplay)}
          </button>
        )}
        {status === "playing" && (
          <button
            onClick={handleShowAnswer}
            className="ghost-btn"
            style={{ ...ghostBtnStyle, borderColor: COLORS.gold, color: COLORS.gold }}
          >
            {t("play_show_answer", meaningDisplay)}
          </button>
        )}
        {status !== "playing" && (
          <button onClick={buildRound} className="seal-btn" style={sealBtnStyle}>
            {t("play_next_char", meaningDisplay)}
          </button>
        )}
      </div>
    </div>
  );
}

const ghostBtnStyle = {
  background: "transparent",
  border: `1px solid ${COLORS.hairline}`,
  color: COLORS.inkSoft,
  borderRadius: 9,
  padding: "9px 17px",
  fontSize: 13.5,
  fontWeight: 500,
  cursor: "pointer",
};

const sealBtnStyle = {
  background: COLORS.seal,
  border: "none",
  color: "#FBF9EF",
  borderRadius: 9,
  padding: "10px 22px",
  fontSize: 14,
  fontWeight: 600,
  cursor: "pointer",
  letterSpacing: 0.2,
};

/* ---------- Simplified SM-2 (the algorithm behind Anki). Given a card's
   current progress (or null for a never-reviewed card) and a rating, works
   out the new ease/interval/repetitions and the date it's next due. ---------- */
function updateSM2(progress, rating) {
  let ease = progress ? progress.ease_factor : 2.5;
  let interval = progress ? progress.interval_days : 0;
  let reps = progress ? progress.repetitions : 0;

  if (rating === "again") {
    reps = 0;
    interval = 1;
    ease = Math.max(1.3, ease - 0.2);
  } else if (rating === "hard") {
    ease = Math.max(1.3, ease - 0.15);
    interval = reps === 0 ? 1 : Math.max(1, Math.round(interval * 1.2));
    reps += 1;
  } else if (rating === "good") {
    if (reps === 0) interval = 1;
    else if (reps === 1) interval = 6;
    else interval = Math.max(1, Math.round(interval * ease));
    reps += 1;
  } else if (rating === "easy") {
    ease = ease + 0.15;
    if (reps === 0) interval = 4;
    else if (reps === 1) interval = 8;
    else interval = Math.max(1, Math.round(interval * ease * 1.3));
    reps += 1;
  }

  const due = new Date();
  due.setDate(due.getDate() + interval);

  return {
    ease_factor: ease,
    interval_days: interval,
    repetitions: reps,
    due_date: due.toISOString().slice(0, 10),
  };
}

/* ================= FLASHCARDS TAB ================= */
function FlashcardsTab({ userId, characterList, wordList, bushouList, decks, isAdmin, checkListAccess, onRequireAuth, onViewPremium, meaningDisplay, goToAddNewCards }) {
  const [contentType, setContentType] = useState("words"); // words = characters+words, radicals = bushou, deck = a saved deck
  const [selectedList, setSelectedList] = useState("Tất cả");
  const [selectedDeckId, setSelectedDeckId] = useState(decks && decks[0] ? decks[0].id : "");
  const [lockedListName, setLockedListName] = useState(null);
  const [progressMap, setProgressMap] = useState(null); // null = loading
  const [queue, setQueue] = useState([]);
  const [current, setCurrent] = useState(null);
  const [flipped, setFlipped] = useState(false);
  const [sessionActive, setSessionActive] = useState(false);
  const [sessionStats, setSessionStats] = useState({ reviewed: 0, again: 0 });

  const selectedDeck = (decks || []).find((d) => d.id === selectedDeckId) || null;

  const allLists = useMemo(() => {
    const set = new Set();
    if (contentType === "radicals") {
      (bushouList || []).forEach((b) => (b.lists || []).forEach((l) => set.add(l.trim())));
      return Array.from(set).sort((a, b) => {
        const numA = parseInt(a, 10);
        const numB = parseInt(b, 10);
        if (!isNaN(numA) && !isNaN(numB)) return numA - numB;
        if (!isNaN(numA)) return -1;
        if (!isNaN(numB)) return 1;
        return a.localeCompare(b, "vi");
      });
    }
    characterList.forEach((c) => getLists(c).forEach((l) => set.add(l.trim())));
    wordList.forEach((w) => (w.lists || []).forEach((l) => set.add(l.trim())));
    return Array.from(set).sort((a, b) => a.localeCompare(b, "vi"));
  }, [characterList, wordList, bushouList, contentType]);

  useEffect(() => {
    if (!userId) {
      setProgressMap(new Map());
      return;
    }
    let cancelled = false;
    (async () => {
      const { data, error } = await supabase.from("flashcard_progress").select("*").eq("user_id", userId);
      if (cancelled) return;
      const map = new Map();
      if (!error) (data || []).forEach((row) => map.set(`${row.card_type}:${row.card_key}`, row));
      setProgressMap(map);
    })();
    return () => {
      cancelled = true;
    };
  }, [userId]);

  // Today's due count for the currently selected list, shown before
  // starting a session so the person knows what they're in for.
  const dueCount = useMemo(() => {
    if (!progressMap) return 0;
    const todayStr = new Date().toISOString().slice(0, 10);
    let count = 0;
    if (contentType === "deck") {
      resolveDeckItems(selectedDeck, { characterList, wordList, bushouList }).forEach((item) => {
        const p = progressMap.get(`${item.type}:${item.key}`);
        if (!p || p.due_date <= todayStr) count += 1;
      });
      return count;
    }
    if (contentType === "radicals") {
      (bushouList || []).forEach((b) => {
        if (selectedList !== "Tất cả" && !(b.lists || []).some((l) => l.trim() === selectedList)) return;
        const p = progressMap.get(`bushou:${b.char}`);
        if (!p || p.due_date <= todayStr) count += 1;
      });
      return count;
    }
    characterList.forEach((c) => {
      if (selectedList !== "Tất cả" && !getLists(c).some((l) => l.trim() === selectedList)) return;
      const p = progressMap.get(`char:${c.char}`);
      if (!p || p.due_date <= todayStr) count += 1;
    });
    wordList.forEach((w) => {
      if (selectedList !== "Tất cả" && !(w.lists || []).some((l) => l.trim() === selectedList)) return;
      const p = progressMap.get(`word:${w.word}`);
      if (!p || p.due_date <= todayStr) count += 1;
    });
    return count;
  }, [progressMap, characterList, wordList, bushouList, contentType, selectedList, selectedDeck]);

  function handleContentTypeChange(next) {
    setContentType(next);
    setSelectedList("Tất cả");
  }

  function handleListChange(next) {
    if (!isAdmin && next !== "Tất cả" && checkListAccess && !checkListAccess(next)) {
      setLockedListName(next);
      return;
    }
    setSelectedList(next);
  }

  function startSession() {
    const todayStr = new Date().toISOString().slice(0, 10);
    const cards = [];
    if (contentType === "deck") {
      resolveDeckItems(selectedDeck, { characterList, wordList, bushouList }).forEach((item) => {
        const p = progressMap.get(`${item.type}:${item.key}`);
        if (!p || p.due_date <= todayStr) cards.push({ type: item.type, key: item.key, data: item.data, progress: p || null });
      });
    } else if (contentType === "radicals") {
      (bushouList || []).forEach((b) => {
        if (selectedList !== "Tất cả" && !(b.lists || []).some((l) => l.trim() === selectedList)) return;
        const p = progressMap.get(`bushou:${b.char}`);
        if (!p || p.due_date <= todayStr) cards.push({ type: "bushou", key: b.char, data: b, progress: p || null });
      });
    } else {
      characterList.forEach((c) => {
        if (selectedList !== "Tất cả" && !getLists(c).some((l) => l.trim() === selectedList)) return;
        const p = progressMap.get(`char:${c.char}`);
        if (!p || p.due_date <= todayStr) cards.push({ type: "char", key: c.char, data: c, progress: p || null });
      });
      wordList.forEach((w) => {
        if (selectedList !== "Tất cả" && !(w.lists || []).some((l) => l.trim() === selectedList)) return;
        const p = progressMap.get(`word:${w.word}`);
        if (!p || p.due_date <= todayStr) cards.push({ type: "word", key: w.word, data: w, progress: p || null });
      });
    }
    const shuffled = shuffle(cards);
    setQueue(shuffled.slice(1));
    setCurrent(shuffled[0] || null);
    setFlipped(false);
    setSessionStats({ reviewed: 0, again: 0 });
    setSessionActive(true);
  }

  async function rate(rating) {
    if (!current) return;
    const updated = updateSM2(current.progress, rating);
    const row = {
      user_id: userId,
      card_key: current.key,
      card_type: current.type,
      ease_factor: updated.ease_factor,
      interval_days: updated.interval_days,
      repetitions: updated.repetitions,
      due_date: updated.due_date,
      last_reviewed: new Date().toISOString(),
    };
    setProgressMap((prev) => {
      const next = new Map(prev);
      next.set(`${current.type}:${current.key}`, row);
      return next;
    });
    if (userId) {
      supabase
        .from("flashcard_progress")
        .upsert(row, { onConflict: "user_id,card_key,card_type" })
        .then(({ error }) => {
          if (error) console.error("Could not save flashcard progress:", error);
        });
    }
    // Guests: progress only lives in progressMap for this session -- there's
    // no account to persist it to, same as the rest of the app's guest mode.

    setSessionStats((prev) => ({ reviewed: prev.reviewed + 1, again: prev.again + (rating === "again" ? 1 : 0) }));
    const rest = queue;
    setQueue(rest.slice(1));
    setCurrent(rest[0] || null);
    setFlipped(false);
  }

  function endSession() {
    setSessionActive(false);
    setCurrent(null);
    setQueue([]);
  }

  if (progressMap === null) {
    return <div style={{ textAlign: "center", color: COLORS.inkSoft, padding: 40 }}>{t("loading", meaningDisplay)}</div>;
  }

  return (
    <div style={{ maxWidth: 420, margin: "0 auto" }}>
      {lockedListName && (
        <ListLockedModal listName={lockedListName} onClose={() => setLockedListName(null)} onViewPremium={onViewPremium} />
      )}

      {!sessionActive ? (
        <div style={{ textAlign: "center" }}>
          <div style={{ fontSize: 12.5, fontWeight: 600, color: COLORS.gold, marginBottom: 16, textTransform: "uppercase", letterSpacing: 0.8 }}>
            {t("fc_title", meaningDisplay)}
          </div>

          <div style={{ display: "flex", justifyContent: "center", marginBottom: 14 }}>
            <button
              type="button"
              onClick={goToAddNewCards}
              style={{
                border: `1.5px solid ${COLORS.gold}`,
                borderRadius: 999,
                background: "transparent",
                color: COLORS.gold,
                fontWeight: 700,
                fontSize: 12.5,
                padding: "7px 14px",
                cursor: "pointer",
              }}
            >
              + {t("tab_add", meaningDisplay)}
            </button>
          </div>

          <div style={{ display: "flex", justifyContent: "center", gap: 6, marginBottom: 14 }}>
            <button
              type="button"
              onClick={() => handleContentTypeChange("words")}
              style={{
                padding: "6px 14px",
                borderRadius: 999,
                border: `1.5px solid ${contentType === "words" ? COLORS.seal : COLORS.hairline}`,
                background: contentType === "words" ? COLORS.seal : "transparent",
                color: contentType === "words" ? "#FBF9EF" : COLORS.inkSoft,
                fontSize: 12.5,
                fontWeight: 600,
                cursor: "pointer",
              }}
            >
              {t("fc_content_words", meaningDisplay)}
            </button>
            <button
              type="button"
              onClick={() => handleContentTypeChange("radicals")}
              style={{
                padding: "6px 14px",
                borderRadius: 999,
                border: `1.5px solid ${contentType === "radicals" ? COLORS.seal : COLORS.hairline}`,
                background: contentType === "radicals" ? COLORS.seal : "transparent",
                color: contentType === "radicals" ? "#FBF9EF" : COLORS.inkSoft,
                fontSize: 12.5,
                fontWeight: 600,
                cursor: "pointer",
              }}
            >
              {t("fc_content_radicals", meaningDisplay)}
            </button>
            {decks && decks.length > 0 && (
              <button
                type="button"
                onClick={() => handleContentTypeChange("deck")}
                style={{
                  padding: "6px 14px",
                  borderRadius: 999,
                  border: `1.5px solid ${contentType === "deck" ? COLORS.seal : COLORS.hairline}`,
                  background: contentType === "deck" ? COLORS.seal : "transparent",
                  color: contentType === "deck" ? "#FBF9EF" : COLORS.inkSoft,
                  fontSize: 12.5,
                  fontWeight: 600,
                  cursor: "pointer",
                }}
              >
                📦 {t("fc_content_deck", meaningDisplay)}
              </button>
            )}
          </div>

          {contentType === "deck" ? (
            <select
              value={selectedDeckId}
              onChange={(e) => setSelectedDeckId(e.target.value)}
              style={{ ...selectStyle, width: 260, textAlign: "center", display: "inline-block", marginBottom: 16 }}
            >
              {(decks || []).map((d) => (
                <option key={d.id} value={d.id} style={{ background: COLORS.chipBg, color: COLORS.ink, fontWeight: 700 }}>
                  {d.name}
                </option>
              ))}
            </select>
          ) : (
            <select
              value={selectedList}
              onChange={(e) => handleListChange(e.target.value)}
              style={{ ...selectStyle, width: 260, textAlign: "center", display: "inline-block", marginBottom: 16 }}
            >
              <option value="Tất cả" style={{ background: COLORS.chipBg, color: COLORS.ink, fontWeight: 700 }}>{t("fc_all_lists", meaningDisplay)}</option>
              {allLists.map((l) => (
                <option key={l} value={l} style={{ background: COLORS.chipBg, color: COLORS.ink, fontWeight: 700 }}>
                  {!isAdmin && checkListAccess && !checkListAccess(l) ? `🔒 ${l}` : displayListName(l, meaningDisplay)}
                </option>
              ))}
            </select>
          )}

          <div style={{ fontSize: 14, color: COLORS.inkSoft, marginBottom: 20 }}>
            {t("fc_due_today", meaningDisplay, dueCount)}
          </div>

          <button
            type="button"
            onClick={startSession}
            disabled={dueCount === 0}
            className="seal-btn"
            style={{ ...sealBtnStyle, padding: "10px 26px", fontSize: 14, opacity: dueCount === 0 ? 0.5 : 1 }}
          >
            {t("fc_start", meaningDisplay)}
          </button>
        </div>
      ) : current ? (
        <div>
          <div style={{ fontSize: 12, color: COLORS.inkSoft, textAlign: "center", marginBottom: 14 }}>
            {t("fc_progress", meaningDisplay, sessionStats.reviewed, queue.length + 1)}
          </div>

          <div
            onClick={() => setFlipped((f) => !f)}
            style={{
              background: COLORS.card,
              border: `1px solid ${COLORS.hairline}`,
              borderRadius: 14,
              padding: "40px 24px",
              minHeight: 220,
              display: "flex",
              flexDirection: "column",
              alignItems: "center",
              justifyContent: "center",
              cursor: "pointer",
              marginBottom: 16,
              textAlign: "center",
            }}
          >
            {!flipped ? (
              <div style={{ fontFamily: "'Noto Serif SC', 'STKaiti', 'Kaiti SC', serif", fontSize: current.type === "word" ? 44 : 64, color: COLORS.ink }}>
                {current.key}
              </div>
            ) : (
              <div style={{ width: "100%" }}>
                <div style={{ fontFamily: "'Noto Serif SC', 'STKaiti', 'Kaiti SC', serif", fontSize: 34, color: COLORS.ink, marginBottom: 10 }}>
                  {current.key}
                </div>
                <div style={{ fontSize: 16, color: COLORS.sealDark, marginBottom: 6 }}>{current.data.pinyin}</div>
                <div style={{ marginBottom: 4 }}>
                  <MeaningBoxes meaning={current.data.meaning} meaningVi={current.data.meaning_vi} meaningDisplay={meaningDisplay} large />
                </div>
                {current.data.sv && meaningDisplay !== "en" && (
                  <div style={{ fontSize: 13.5, color: COLORS.bamboo, fontWeight: 600, marginTop: 4 }}>HV: {current.data.sv}</div>
                )}
              </div>
            )}
          </div>

          {!flipped ? (
            <div style={{ textAlign: "center" }}>
              <button type="button" onClick={() => setFlipped(true)} className="ghost-btn" style={{ ...ghostBtnStyle, padding: "8px 22px", fontSize: 13 }}>
                {t("fc_flip", meaningDisplay)}
              </button>
            </div>
          ) : (
            <div style={{ display: "grid", gridTemplateColumns: "repeat(4, 1fr)", gap: 6 }}>
              <button type="button" onClick={() => rate("again")} style={{ ...ratingBtnStyle, borderColor: COLORS.error, color: COLORS.error }}>
                {t("fc_again", meaningDisplay)}
              </button>
              <button type="button" onClick={() => rate("hard")} style={{ ...ratingBtnStyle, borderColor: COLORS.gold, color: COLORS.gold }}>
                {t("fc_hard", meaningDisplay)}
              </button>
              <button type="button" onClick={() => rate("good")} style={{ ...ratingBtnStyle, borderColor: COLORS.seal, color: COLORS.seal }}>
                {t("fc_good", meaningDisplay)}
              </button>
              <button type="button" onClick={() => rate("easy")} style={{ ...ratingBtnStyle, borderColor: COLORS.bamboo, color: COLORS.bamboo }}>
                {t("fc_easy", meaningDisplay)}
              </button>
            </div>
          )}

          <div style={{ textAlign: "center", marginTop: 16 }}>
            <button type="button" onClick={endSession} className="ghost-btn" style={{ ...ghostBtnStyle, padding: "6px 16px", fontSize: 12 }}>
              {t("fc_end", meaningDisplay)}
            </button>
          </div>
        </div>
      ) : (
        <div style={{ textAlign: "center" }}>
          <div style={{ fontSize: 20, fontWeight: 700, color: COLORS.bamboo, marginBottom: 8 }}>{t("fc_complete_title", meaningDisplay)}</div>
          <div style={{ fontSize: 14, color: COLORS.inkSoft, marginBottom: 20 }}>
            {t("fc_complete_summary", meaningDisplay, sessionStats.reviewed, sessionStats.again)}
          </div>
          <button type="button" onClick={endSession} className="seal-btn" style={{ ...sealBtnStyle, padding: "10px 26px", fontSize: 14 }}>
            {t("fc_done", meaningDisplay)}
          </button>
        </div>
      )}
    </div>
  );
}

const ratingBtnStyle = {
  padding: "10px 4px",
  borderRadius: 11,
  border: "1.5px solid",
  background: "transparent",
  fontSize: 12,
  fontWeight: 700,
  cursor: "pointer",
};

/* ================= WRITING PRACTICE TAB =================
   Guided stroke tracing using HanziWriter's built-in quiz() mode — it
   already handles both mouse and touch input identically, and already
   judges each stroke against the real reference data (the same data the
   stroke-order animation uses), so there's no recognition logic to build
   here. No login required and no API cost: this never calls our lookup
   functions, HanziWriter fetches character stroke data from its own
   public source. ---------- */
function WritingPracticeTab({ characterList, bushouList, decks, isAdmin, checkListAccess, onViewPremium, meaningDisplay, goToAddNewCards }) {
  const [contentType, setContentType] = useState("chars"); // chars | radicals | deck
  const [selectedList, setSelectedList] = useState("Tất cả");
  const [selectedDeckId, setSelectedDeckId] = useState(decks && decks[0] ? decks[0].id : "");
  const [lockedListName, setLockedListName] = useState(null);
  const [searchQuery, setSearchQuery] = useState("");
  const [previewPage, setPreviewPage] = useState(0);
  const [brushSize, setBrushSize] = useState("normal"); // thin | normal | thick
  const [queue, setQueue] = useState([]);
  const [currentIndex, setCurrentIndex] = useState(0);
  const [sessionActive, setSessionActive] = useState(false);
  const [mistakes, setMistakes] = useState(0);
  // "loading" | "demo" | "practicing" | "done-char" | "dots" | "recall" | "error"
  const [status, setStatus] = useState("idle");
  const [revealOn, setRevealOn] = useState(false);
  const targetRef = useRef(null);
  const writerRef = useRef(null);
  const recallCanvasRef = useRef(null);
  const recallStrokesRef = useRef([]); // array of strokes, each a list of {x,y} points
  const isDrawingRef = useRef(false);

  // Dot-connecting step (between guided tracing and recall): shows each
  // stroke's real start/end point (from HanziWriter's own raw median
  // data) and checks whether the user's drawn line starts and ends near
  // enough to those two points.
  const [dotCharData, setDotCharData] = useState(null); // { medians, strokes } | null
  const [dotStrokeIndex, setDotStrokeIndex] = useState(0);
  const [completedDotStrokes, setCompletedDotStrokes] = useState([]);
  const [dotFeedback, setDotFeedback] = useState(null); // null | "correct" | "wrong"
  const dotsCanvasRef = useRef(null);
  const dotDrawStartRef = useRef(null);
  const dotCurrentPathRef = useRef([]);

  const RECALL_INK_COLOR = "#2456A6";
  const BRUSH_WIDTHS = { thin: 23.4, normal: 39, thick: 62.4 };
  // HanziWriter's drawingWidth is specified in its own internal SVG
  // coordinate space and gets scaled down to fit the display size, so the
  // same raw number looks much thinner there than it would as a plain
  // canvas lineWidth (which is 1 unit = 1 actual pixel, no scaling). This
  // is the recall canvas's own scale, tuned to visually match the guided
  // tracing pen at the same "Cỡ bút" selection.
  const RECALL_BRUSH_WIDTHS = { thin: 5.85, normal: 9.75, thick: 15.6 };
  const PREVIEW_PER_PAGE = 24; // ~8 columns x 3 rows at this layout's width

  const allLists = useMemo(() => {
    const set = new Set();
    if (contentType === "radicals") {
      (bushouList || []).forEach((b) => (b.lists || []).forEach((l) => set.add(l.trim())));
      return Array.from(set).sort((a, b) => {
        const numA = parseInt(a, 10);
        const numB = parseInt(b, 10);
        if (!isNaN(numA) && !isNaN(numB)) return numA - numB;
        if (!isNaN(numA)) return -1;
        if (!isNaN(numB)) return 1;
        return a.localeCompare(b, "vi");
      });
    }
    characterList.forEach((c) => getLists(c).forEach((l) => set.add(l.trim())));
    return Array.from(set).sort((a, b) => a.localeCompare(b, "vi"));
  }, [characterList, bushouList, contentType]);

  const selectedDeck = (decks || []).find((d) => d.id === selectedDeckId) || null;

  // The pool of practiceable items for the current content-type selection --
  // used by both the preview grid and starting a session, so they always
  // agree on what's actually available.
  function getPool() {
    if (contentType === "deck") {
      return resolveDeckItems(selectedDeck, { characterList, wordList: [], bushouList }).map((item) => item.data);
    }
    if (contentType === "radicals") {
      return (bushouList || []).filter((b) => selectedList === "Tất cả" || (b.lists || []).some((l) => l.trim() === selectedList));
    }
    return characterList.filter((c) => selectedList === "Tất cả" || getLists(c).some((l) => l.trim() === selectedList));
  }

  const previewChars = useMemo(() => getPool(), [characterList, bushouList, decks, contentType, selectedList, selectedDeckId]);

  const previewTotalPages = Math.max(1, Math.ceil(previewChars.length / PREVIEW_PER_PAGE));
  const previewPageItems = previewChars.slice(previewPage * PREVIEW_PER_PAGE, (previewPage + 1) * PREVIEW_PER_PAGE);

  const searchResults = useMemo(() => {
    const q = searchQuery.trim().toLowerCase();
    if (!q) return [];
    const source = contentType === "radicals" ? bushouList || [] : characterList;
    return source
      .filter((c) => c.char.includes(q) || c.pinyin.toLowerCase().includes(q) || c.meaning.toLowerCase().includes(q))
      .slice(0, 8);
  }, [searchQuery, characterList, bushouList, contentType]);

  function handleContentTypeChange(next) {
    setContentType(next);
    setSelectedList("Tất cả");
    setPreviewPage(0);
  }

  function handleListChange(next) {
    if (!isAdmin && next !== "Tất cả" && checkListAccess && !checkListAccess(next)) {
      setLockedListName(next);
      return;
    }
    setSelectedList(next);
    setPreviewPage(0);
  }

  // Same access check as handleListChange, but usable mid-session: swaps
  // in a fresh shuffled queue from the new list without returning to the
  // start screen.
  function switchListMidSession(next) {
    if (!isAdmin && next !== "Tất cả" && checkListAccess && !checkListAccess(next)) {
      setLockedListName(next);
      return;
    }
    setSelectedList(next);
    const items =
      contentType === "radicals"
        ? (bushouList || []).filter((b) => next === "Tất cả" || (b.lists || []).some((l) => l.trim() === next))
        : characterList.filter((c) => next === "Tất cả" || getLists(c).some((l) => l.trim() === next));
    if (items.length === 0) return;
    setStatus("loading");
    setRevealOn(false);
    setQueue(shuffle(items));
    setCurrentIndex(0);
  }

  function startSession() {
    const items = getPool();
    if (items.length === 0) return;
    setStatus("loading");
    setRevealOn(false);
    setQueue(shuffle(items));
    setCurrentIndex(0);
    setSessionActive(true);
  }

  function startSingleChar(c) {
    // Start with this character, then continue through the rest of the
    // currently selected list (shuffled) -- so "Chữ tiếp theo" has
    // somewhere to go instead of ending the session immediately.
    const rest = shuffle(previewChars.filter((x) => x.char !== c.char));
    setStatus("loading");
    setRevealOn(false);
    setQueue([c, ...rest]);
    setCurrentIndex(0);
    setSessionActive(true);
    setSearchQuery("");
  }

  function endSession() {
    setSessionActive(false);
    setStatus("idle");
    setRevealOn(false);
    setQueue([]);
    setCurrentIndex(0);
  }

  const current = sessionActive ? queue[currentIndex] : null;

  function playDemo(writer) {
    setStatus("demo");
    writer.animateCharacter();
  }

  function runQuiz(writer) {
    setMistakes(0);
    setStatus("practicing");
    writer.quiz({
      showHintAfterMisses: 3,
      onMistake: () => setMistakes((m) => m + 1),
      onComplete: () => setStatus("done-char"),
    });
  }

  useEffect(() => {
    if (!sessionActive || !current || !targetRef.current) return;
    let cancelled = false;
    setStatus("loading");
    targetRef.current.innerHTML = "";
    try {
      const writer = HanziWriter.create(targetRef.current, current.char, {
        width: 280,
        height: 280,
        padding: 12,
        showOutline: true,
        showCharacter: false,
        highlightOnComplete: true,
        highlightCompleteColor: COLORS.bamboo,
        strokeColor: COLORS.ink,
        outlineColor: COLORS.grid,
        drawingWidth: BRUSH_WIDTHS[brushSize],
        onLoadCharDataSuccess: () => {
          if (cancelled) return;
          writerRef.current = writer;
          playDemo(writer);
        },
        onLoadCharDataError: () => {
          if (!cancelled) setStatus("error");
        },
      });
    } catch (e) {
      console.error("Writing practice failed to load:", e);
      setStatus("error");
    }
    return () => {
      cancelled = true;
    };
    // eslint-disable-next-line react-hooks/exhaustive-deps
  }, [sessionActive, currentIndex, current, brushSize]);

  function replayDemo() {
    if (writerRef.current) playDemo(writerRef.current);
  }

  function beginWriting() {
    if (writerRef.current) runQuiz(writerRef.current);
  }

  function retryChar() {
    if (writerRef.current) runQuiz(writerRef.current);
  }

  function nextChar() {
    if (currentIndex + 1 >= queue.length) {
      endSession();
      return;
    }
    // Explicitly step out of "recall" status (and its reveal toggle) here,
    // in the same update as the index change -- otherwise the writing
    // canvas and the HanziWriter container, which are mutually exclusive
    // in the layout based on status, briefly (or permanently) disagree
    // about which one should exist, and the new character's demo never
    // gets a real container to initialize into.
    setStatus("loading");
    setRevealOn(false);
    setCurrentIndex((i) => i + 1);
  }

  function enterDotConnectMode() {
    setDotCharData(null);
    setDotStrokeIndex(0);
    setCompletedDotStrokes([]);
    setDotFeedback(null);
    setRevealOn(false);
    setStatus("dots");
  }

  function enterRecallMode() {
    recallStrokesRef.current = [];
    setRevealOn(false);
    setStatus("recall");
  }

  // Explicitly wipes the actual canvas pixels on entering recall mode.
  // recallStrokesRef being reset (above) only clears the stroke-history
  // data -- if the browser happens to reuse the same underlying <canvas>
  // element that the dot-connecting step was just drawing on (React can
  // do this since both sit in the same layout position), the old pixels
  // physically remain until something explicitly clears them.
  useEffect(() => {
    if (status !== "recall") return;
    const canvas = recallCanvasRef.current;
    if (!canvas) return;
    canvas.getContext("2d").clearRect(0, 0, canvas.width, canvas.height);
  }, [status, current]);

  // Loads the raw stroke data (medians = each stroke's start/end/path
  // points) for the dot-connecting step. Uses HanziWriter's own loader,
  // same data source as everything else -- no new dependency, no cost.
  useEffect(() => {
    if (status !== "dots" || !current) return;
    let cancelled = false;
    HanziWriter.loadCharacterData(current.char)
      .then((data) => {
        if (cancelled) return;
        if (!data || !Array.isArray(data.medians) || data.medians.length === 0) {
          console.error("Stroke data for dot practice is missing medians:", current.char, data);
          setDotFeedback("load-error");
          return;
        }
        setDotCharData(data);
      })
      .catch((e) => {
        if (!cancelled) {
          console.error("Could not load stroke data for dot practice:", e);
          setDotFeedback("load-error");
        }
      });
    return () => {
      cancelled = true;
    };
  }, [status, current]);

  // This data format lays every character's raw stroke coordinates out on
  // a fixed 1024x1024 grid where, unusually, the origin is offset and the
  // y-axis runs the opposite way from normal screen coordinates: (0, 900)
  // is the top-left corner and (1024, -124) is the bottom-right. This is a
  // documented constant of the data format itself (same for every
  // character), not something to detect per-character -- so rather than
  // parse a transform string of uncertain exact format, compute the
  // conversion directly from these known values.
  const dotTransform = useMemo(() => {
    if (!dotCharData) return null;
    const padding = 12;
    const viewSize = 1024;
    const scale = (280 - padding * 2) / viewSize;
    return (pt) => ({
      x: padding + pt[0] * scale,
      y: padding + (900 - pt[1]) * scale,
    });
  }, [dotCharData]);

  // Same transform as dotTransform above, expressed as an SVG matrix so
  // the real stroke shapes (not a font rendering of the character) can be
  // drawn as the reveal reference -- guaranteed to align exactly with the
  // dots since both come from the same source data and the same math.
  const dotSvgMatrix = useMemo(() => {
    if (!dotCharData) return null;
    const padding = 12;
    const scale = (280 - padding * 2) / 1024;
    return `matrix(${scale},0,0,${-scale},${padding},${padding + 900 * scale})`;
  }, [dotCharData]);

  const dotTotalStrokes = dotCharData && dotCharData.medians ? dotCharData.medians.length : 0;
  const dotCurrentMedian = dotCharData && dotCharData.medians ? dotCharData.medians[dotStrokeIndex] : null;
  const dotStartPoint = dotCurrentMedian && dotTransform ? dotTransform(dotCurrentMedian[0]) : null;
  const dotEndPoint = dotCurrentMedian && dotTransform ? dotTransform(dotCurrentMedian[dotCurrentMedian.length - 1]) : null;

  function redrawDotsCanvas(completedList) {
    const canvas = dotsCanvasRef.current;
    if (!canvas) return;
    const ctx = canvas.getContext("2d");
    ctx.clearRect(0, 0, canvas.width, canvas.height);
    const baseWidth = RECALL_BRUSH_WIDTHS[brushSize];
    completedList.forEach((entry) => {
      const path = entry.path;
      if (!path || path.length < 2) return;
      const widths = computeInkWidths(path, baseWidth);
      drawVariableWidthPath(ctx, path, widths, COLORS.bamboo);
    });
  }

  useEffect(() => {
    redrawDotsCanvas(completedDotStrokes);
    // eslint-disable-next-line react-hooks/exhaustive-deps
  }, [dotCharData, dotStrokeIndex, brushSize]);

  function startDotDrawing(e) {
    e.preventDefault();
    isDrawingRef.current = true;
    const pt = getPointForCanvas(dotsCanvasRef, e.nativeEvent);
    dotDrawStartRef.current = pt;
    dotCurrentPathRef.current = [pt];
  }

  function continueDotDrawing(e) {
    if (!isDrawingRef.current) return;
    e.preventDefault();
    const pt = getPointForCanvas(dotsCanvasRef, e.nativeEvent);
    const path = dotCurrentPathRef.current;
    path.push(pt);
    if (path.length < 2) return;
    // Redraw completed strokes, then the current in-progress path on top --
    // needed (rather than drawing just the newest segment) so the brush's
    // speed-based width can respond smoothly along the whole stroke.
    const canvas = dotsCanvasRef.current;
    if (!canvas) return;
    const ctx = canvas.getContext("2d");
    redrawDotsCanvas(completedDotStrokes);
    const baseWidth = RECALL_BRUSH_WIDTHS[brushSize];
    const widths = computeInkWidths(path, baseWidth, { taperEnd: false });
    drawVariableWidthPath(ctx, path, widths, RECALL_INK_COLOR);
  }

  function endDotDrawing() {
    if (!isDrawingRef.current) return;
    isDrawingRef.current = false;
    const path = dotCurrentPathRef.current;
    dotCurrentPathRef.current = [];
    dotDrawStartRef.current = null;
    if (path.length < 2 || !dotStartPoint || !dotEndPoint) return;
    const startPt = path[0];
    const endPt = path[path.length - 1];

    // Only the actual start and end of what was drawn matter here -- the
    // path in between can be any shape (straight, curved, zigzagged), it
    // isn't checked at all.
    const THRESHOLD = 30; // pixels -- generous enough for finger/mouse imprecision
    const distToStart = Math.hypot(startPt.x - dotStartPoint.x, startPt.y - dotStartPoint.y);
    const distToEnd = Math.hypot(endPt.x - dotEndPoint.x, endPt.y - dotEndPoint.y);

    if (distToStart <= THRESHOLD && distToEnd <= THRESHOLD) {
      setDotFeedback("correct");
      const nextCompleted = [...completedDotStrokes, { strokeIndex: dotStrokeIndex, path }];
      setCompletedDotStrokes(nextCompleted);
      if (dotStrokeIndex + 1 >= dotTotalStrokes) {
        // All strokes connected -- stop here and let the person look at
        // the finished character before moving on; they advance manually
        // via the "Viết từ trí nhớ" button below.
        setTimeout(() => setDotFeedback(null), 400);
      } else {
        setTimeout(() => {
          setDotFeedback(null);
          setDotStrokeIndex((i) => i + 1);
        }, 400);
      }
    } else {
      setDotFeedback("wrong");
      redrawDotsCanvas(completedDotStrokes);
      setTimeout(() => setDotFeedback(null), 500);
    }
  }

  function undoLastDotStroke() {
    if (completedDotStrokes.length === 0) return;
    const last = completedDotStrokes[completedDotStrokes.length - 1];
    const remaining = completedDotStrokes.slice(0, -1);
    setCompletedDotStrokes(remaining);
    setDotStrokeIndex(last.strokeIndex);
    setDotFeedback(null);
    redrawDotsCanvas(remaining);
  }

  function getPointForCanvas(canvasRef, nativeEvent) {
    const canvas = canvasRef.current;
    const rect = canvas.getBoundingClientRect();
    const scaleX = canvas.width / rect.width;
    const scaleY = canvas.height / rect.height;
    return {
      x: (nativeEvent.clientX - rect.left) * scaleX,
      y: (nativeEvent.clientY - rect.top) * scaleY,
      t: nativeEvent.timeStamp || performance.now(),
      pressure: nativeEvent.pressure,
      pointerType: nativeEvent.pointerType,
    };
  }

  // Subtle "natural ink" brush: width varies gently with drawing speed
  // (slower = a bit thicker, faster = a bit thinner), uses real pressure
  // instead when a stylus reports it, and tapers thin at the very start
  // (and, once a stroke is finalized, the very end) of each stroke --
  // approximating the tapered, speed-responsive feel of a natural ink
  // brush rather than a flat, uniform marker line.
  function computeInkWidths(path, baseWidth, { taperEnd = true } = {}) {
    const n = path.length;
    if (n === 0) return [];
    const speeds = new Array(n).fill(0);
    for (let i = 1; i < n; i++) {
      const dist = Math.hypot(path[i].x - path[i - 1].x, path[i].y - path[i - 1].y);
      const dt = Math.max(1, (path[i].t || 0) - (path[i - 1].t || 0));
      speeds[i] = dist / dt;
    }
    const maxSpeed = Math.max(0.02, ...speeds);
    const taperLen = Math.min(6, Math.floor(n / 2));
    const widths = new Array(n);
    for (let i = 0; i < n; i++) {
      const pt = path[i];
      let factor;
      if (pt.pointerType === "pen" && typeof pt.pressure === "number" && pt.pressure > 0) {
        factor = 0.6 + pt.pressure * 0.8; // real stylus pressure -> 0.6x-1.4x
      } else {
        const speedNorm = speeds[i] / maxSpeed;
        factor = 1.35 - speedNorm * 0.7; // simulated from speed -> 0.65x-1.35x
      }
      let taper = 1;
      if (i < taperLen) taper = 0.35 + 0.65 * (i / taperLen);
      else if (taperEnd && i > n - 1 - taperLen) taper = 0.35 + 0.65 * ((n - 1 - i) / taperLen);
      // A flat minimum width keeps thin-brush strokes from vanishing to a
      // sliver at the taper's thinnest point, while still tapering visibly.
      widths[i] = Math.max(1.2, baseWidth * factor * taper);
    }
    return widths;
  }

  function drawVariableWidthPath(ctx, path, widths, color) {
    if (path.length < 2) return;
    ctx.strokeStyle = color;
    ctx.lineCap = "round";
    ctx.lineJoin = "round";
    for (let i = 1; i < path.length; i++) {
      ctx.lineWidth = (widths[i - 1] + widths[i]) / 2;
      ctx.beginPath();
      ctx.moveTo(path[i - 1].x, path[i - 1].y);
      ctx.lineTo(path[i].x, path[i].y);
      ctx.stroke();
    }
  }

  function getRecallPoint(nativeEvent) {
    return getPointForCanvas(recallCanvasRef, nativeEvent);
  }

  function redrawRecallCanvas() {
    const canvas = recallCanvasRef.current;
    if (!canvas) return;
    const ctx = canvas.getContext("2d");
    ctx.clearRect(0, 0, canvas.width, canvas.height);
    const baseWidth = RECALL_BRUSH_WIDTHS[brushSize];
    recallStrokesRef.current.forEach((stroke) => {
      if (stroke.length < 2) return;
      const widths = computeInkWidths(stroke, baseWidth);
      drawVariableWidthPath(ctx, stroke, widths, RECALL_INK_COLOR);
    });
  }

  function startRecallDrawing(e) {
    e.preventDefault();
    isDrawingRef.current = true;
    recallStrokesRef.current.push([getRecallPoint(e.nativeEvent)]);
  }

  function continueRecallDrawing(e) {
    if (!isDrawingRef.current) return;
    e.preventDefault();
    const strokes = recallStrokesRef.current;
    strokes[strokes.length - 1].push(getRecallPoint(e.nativeEvent));
    // Full redraw (rather than just the newest segment) so the brush's
    // speed-based width responds smoothly along the whole stroke.
    redrawRecallCanvas();
  }

  function endRecallDrawing() {
    isDrawingRef.current = false;
  }

  function undoLastStroke() {
    recallStrokesRef.current.pop();
    redrawRecallCanvas();
  }

  function clearRecallCanvas() {
    recallStrokesRef.current = [];
    redrawRecallCanvas();
  }

  const gridSize = 280;
  const mid = gridSize / 2;
  const inset = gridSize * 0.024;
  const far = gridSize - inset;

  return (
    <div style={{ maxWidth: 420, margin: "0 auto" }}>
      {lockedListName && (
        <ListLockedModal listName={lockedListName} onClose={() => setLockedListName(null)} onViewPremium={onViewPremium} />
      )}

      {!sessionActive ? (
        <div style={{ textAlign: "center" }}>
          <div style={{ fontSize: 12.5, fontWeight: 600, color: COLORS.gold, marginBottom: 16, textTransform: "uppercase", letterSpacing: 0.8 }}>
            {t("wp_title", meaningDisplay)}
          </div>

          <div style={{ display: "flex", justifyContent: "center", marginBottom: 14 }}>
            <button
              type="button"
              onClick={goToAddNewCards}
              style={{
                border: `1.5px solid ${COLORS.gold}`,
                borderRadius: 999,
                background: "transparent",
                color: COLORS.gold,
                fontWeight: 700,
                fontSize: 12.5,
                padding: "7px 14px",
                cursor: "pointer",
              }}
            >
              + {t("tab_add", meaningDisplay)}
            </button>
          </div>

          <div style={{ display: "flex", justifyContent: "center", gap: 6, marginBottom: 16 }}>
            <button
              type="button"
              onClick={() => handleContentTypeChange("chars")}
              style={{
                padding: "6px 14px",
                borderRadius: 999,
                border: `1.5px solid ${contentType === "chars" ? COLORS.seal : COLORS.hairline}`,
                background: contentType === "chars" ? COLORS.seal : "transparent",
                color: contentType === "chars" ? "#FBF9EF" : COLORS.inkSoft,
                fontSize: 12.5,
                fontWeight: 600,
                cursor: "pointer",
              }}
            >
              {t("fc_content_words", meaningDisplay)}
            </button>
            <button
              type="button"
              onClick={() => handleContentTypeChange("radicals")}
              style={{
                padding: "6px 14px",
                borderRadius: 999,
                border: `1.5px solid ${contentType === "radicals" ? COLORS.seal : COLORS.hairline}`,
                background: contentType === "radicals" ? COLORS.seal : "transparent",
                color: contentType === "radicals" ? "#FBF9EF" : COLORS.inkSoft,
                fontSize: 12.5,
                fontWeight: 600,
                cursor: "pointer",
              }}
            >
              {t("fc_content_radicals", meaningDisplay)}
            </button>
            {decks && decks.length > 0 && (
              <button
                type="button"
                onClick={() => handleContentTypeChange("deck")}
                style={{
                  padding: "6px 14px",
                  borderRadius: 999,
                  border: `1.5px solid ${contentType === "deck" ? COLORS.seal : COLORS.hairline}`,
                  background: contentType === "deck" ? COLORS.seal : "transparent",
                  color: contentType === "deck" ? "#FBF9EF" : COLORS.inkSoft,
                  fontSize: 12.5,
                  fontWeight: 600,
                  cursor: "pointer",
                }}
              >
                📦 {t("fc_content_deck", meaningDisplay)}
              </button>
            )}
          </div>

          <div style={{ marginBottom: 14 }}>
            <input
              value={searchQuery}
              onChange={(e) => setSearchQuery(e.target.value)}
              placeholder={t("wp_search_placeholder", meaningDisplay)}
              style={{ ...inputStyle, width: 260, textAlign: "center" }}
            />
            {searchResults.length > 0 && (
              <div style={{ display: "flex", flexWrap: "wrap", justifyContent: "center", gap: 6, marginTop: 10 }}>
                {searchResults.map((c) => (
                  <button
                    key={c.char}
                    type="button"
                    onClick={() => startSingleChar(c)}
                    title={`${c.pinyin} · ${c.meaning}`}
                    style={{
                      fontFamily: "'Noto Serif SC', 'STKaiti', 'Kaiti SC', serif",
                      fontSize: 22,
                      width: 44,
                      height: 44,
                      border: `1px solid ${COLORS.hairline}`,
                      borderRadius: 6,
                      background: COLORS.card,
                      color: COLORS.ink,
                      cursor: "pointer",
                    }}
                  >
                    {c.char}
                  </button>
                ))}
              </div>
            )}
          </div>

          <div style={{ fontSize: 11.5, color: COLORS.inkSoft, marginBottom: 10 }}>{t("wp_or_by_list", meaningDisplay)}</div>

          {contentType === "deck" ? (
            <select
              value={selectedDeckId}
              onChange={(e) => setSelectedDeckId(e.target.value)}
              style={{ ...selectStyle, width: 260, textAlign: "center", display: "inline-block", marginBottom: 16 }}
            >
              {(decks || []).map((d) => (
                <option key={d.id} value={d.id} style={{ background: COLORS.chipBg, color: COLORS.ink, fontWeight: 700 }}>
                  {d.name}
                </option>
              ))}
            </select>
          ) : (
            <select
              value={selectedList}
              onChange={(e) => handleListChange(e.target.value)}
              style={{ ...selectStyle, width: 260, textAlign: "center", display: "inline-block", marginBottom: 16 }}
            >
              <option value="Tất cả" style={{ background: COLORS.chipBg, color: COLORS.ink, fontWeight: 700 }}>{t("wp_all_lists", meaningDisplay)}</option>
              {allLists.map((l) => (
                <option key={l} value={l} style={{ background: COLORS.chipBg, color: COLORS.ink, fontWeight: 700 }}>
                  {!isAdmin && checkListAccess && !checkListAccess(l) ? `🔒 ${l}` : displayListName(l, meaningDisplay)}
                </option>
              ))}
            </select>
          )}

          {previewChars.length > 0 && (
            <div style={{ marginBottom: 20 }}>
              <div style={{ display: "flex", flexWrap: "wrap", justifyContent: "center", gap: 6, marginBottom: 8 }}>
                {previewPageItems.map((c) => (
                  <button
                    key={c.char}
                    type="button"
                    onClick={() => startSingleChar(c)}
                    title={`${c.pinyin} · ${c.meaning}`}
                    style={{
                      fontFamily: "'Noto Serif SC', 'STKaiti', 'Kaiti SC', serif",
                      fontSize: 20,
                      width: 38,
                      height: 38,
                      border: `1px solid ${COLORS.hairline}`,
                      borderRadius: 6,
                      background: COLORS.card,
                      color: COLORS.ink,
                      cursor: "pointer",
                    }}
                  >
                    {c.char}
                  </button>
                ))}
              </div>
              {previewTotalPages > 1 && (
                <div style={{ display: "flex", justifyContent: "center", alignItems: "center", gap: 10 }}>
                  <button
                    type="button"
                    onClick={() => setPreviewPage((p) => Math.max(0, p - 1))}
                    disabled={previewPage === 0}
                    className="ghost-btn"
                    style={{ ...ghostBtnStyle, padding: "4px 10px", fontSize: 11.5, opacity: previewPage === 0 ? 0.4 : 1 }}
                  >
                    {t("wp_prev_page", meaningDisplay)}
                  </button>
                  <span style={{ fontSize: 11.5, color: COLORS.inkSoft }}>
                    {previewPage + 1} / {previewTotalPages}
                  </span>
                  <button
                    type="button"
                    onClick={() => setPreviewPage((p) => Math.min(previewTotalPages - 1, p + 1))}
                    disabled={previewPage >= previewTotalPages - 1}
                    className="ghost-btn"
                    style={{ ...ghostBtnStyle, padding: "4px 10px", fontSize: 11.5, opacity: previewPage >= previewTotalPages - 1 ? 0.4 : 1 }}
                  >
                    {t("wp_next_page", meaningDisplay)}
                  </button>
                </div>
              )}
            </div>
          )}

          <div style={{ fontSize: 12.5, color: COLORS.inkSoft, marginBottom: 20, lineHeight: 1.5 }}>
            {t("wp_description", meaningDisplay)}
          </div>

          <button type="button" onClick={startSession} className="seal-btn" style={{ ...sealBtnStyle, padding: "10px 26px", fontSize: 14 }}>
            {t("wp_start", meaningDisplay)}
          </button>
        </div>
      ) : current ? (
        <div>
          <div style={{ fontSize: 12, color: COLORS.inkSoft, textAlign: "center", marginBottom: 10 }}>
            {currentIndex + 1} / {queue.length}
            {status === "practicing" || status === "done-char" ? ` · ${mistakes} lỗi` : ""}
          </div>

          <div style={{ textAlign: "center", marginBottom: 12 }}>
            <select
              value={selectedList}
              onChange={(e) => switchListMidSession(e.target.value)}
              style={{ ...selectStyle, width: 220, textAlign: "center", display: "inline-block", fontSize: 12 }}
            >
              <option value="Tất cả" style={{ background: COLORS.chipBg, color: COLORS.ink, fontWeight: 700 }}>{t("wp_all_lists", meaningDisplay)}</option>
              {allLists.map((l) => (
                <option key={l} value={l} style={{ background: COLORS.chipBg, color: COLORS.ink, fontWeight: 700 }}>
                  {!isAdmin && checkListAccess && !checkListAccess(l) ? `🔒 ${l}` : displayListName(l, meaningDisplay)}
                </option>
              ))}
            </select>
          </div>

          <div style={{ display: "flex", gap: 6, justifyContent: "center", flexWrap: "wrap", marginBottom: 14 }}>
            {meaningDisplay !== "vi" && (
              <div style={{ padding: "3px 7px", borderRadius: 6, background: COLORS.chipBg, border: `1px solid ${COLORS.hairline}`, textAlign: "center" }}>
                <div style={{ fontSize: 9, color: COLORS.inkSoft, fontWeight: 700, textTransform: "uppercase", letterSpacing: 0.3 }}>EN</div>
                <div style={{ fontSize: 11.5, color: COLORS.ink }}>{current.meaning || "—"}</div>
              </div>
            )}
            {meaningDisplay !== "en" && (
              <div style={{ padding: "3px 7px", borderRadius: 6, background: COLORS.chipBg, border: `1px solid ${COLORS.hairline}`, textAlign: "center" }}>
                <div style={{ fontSize: 9, color: COLORS.inkSoft, fontWeight: 700, textTransform: "uppercase", letterSpacing: 0.3 }}>VI</div>
                <div style={{ fontSize: 11.5, color: current.meaning_vi ? COLORS.ink : COLORS.inkSoft, fontStyle: current.meaning_vi ? "normal" : "italic" }}>
                  {current.meaning_vi || "(chưa dịch)"}
                </div>
              </div>
            )}
            <div style={{ padding: "3px 7px", borderRadius: 6, background: COLORS.chipBg, border: `1px solid ${COLORS.hairline}`, textAlign: "center" }}>
              <div style={{ fontSize: 9, color: COLORS.inkSoft, fontWeight: 700, textTransform: "uppercase", letterSpacing: 0.3 }}>Pinyin</div>
              <div style={{ fontSize: 11.5, color: COLORS.sealDark }}>{current.pinyin}</div>
            </div>
            {meaningDisplay !== "en" && (
              <div style={{ padding: "3px 7px", borderRadius: 6, background: COLORS.chipBg, border: `1px solid ${COLORS.hairline}`, textAlign: "center" }}>
                <div style={{ fontSize: 9, color: COLORS.inkSoft, fontWeight: 700, textTransform: "uppercase", letterSpacing: 0.3 }}>Hán Việt</div>
                <div style={{ fontSize: 11.5, color: COLORS.bamboo }}>{current.sv}</div>
              </div>
            )}
          </div>

          <div style={{ display: "flex", justifyContent: "center", marginBottom: 16 }}>
            {status === "recall" ? (
              <div style={{ width: gridSize, height: gridSize, position: "relative", border: `1px solid ${COLORS.hairline}`, borderRadius: 10 }}>
                <svg width={gridSize} height={gridSize} style={{ position: "absolute", inset: 0, opacity: 0.9 }}>
                  <rect x={0} y={0} width={gridSize} height={gridSize} fill={COLORS.card} />
                  <line x1={mid} y1={inset} x2={mid} y2={far} stroke={COLORS.inkSoft} strokeWidth="1.2" strokeDasharray="4 4" />
                  <line x1={inset} y1={mid} x2={far} y2={mid} stroke={COLORS.inkSoft} strokeWidth="1.2" strokeDasharray="4 4" />
                  <line x1={inset} y1={inset} x2={far} y2={far} stroke={COLORS.inkSoft} strokeWidth="1.2" strokeDasharray="3 5" />
                  <line x1={far} y1={inset} x2={inset} y2={far} stroke={COLORS.inkSoft} strokeWidth="1.2" strokeDasharray="3 5" />
                </svg>
                {revealOn && (
                  <div
                    style={{
                      position: "absolute",
                      inset: 0,
                      display: "flex",
                      alignItems: "center",
                      justifyContent: "center",
                      pointerEvents: "none",
                    }}
                  >
                    <div
                      style={{
                        fontFamily: "'Noto Serif SC', 'STKaiti', 'Kaiti SC', serif",
                        fontSize: gridSize * 0.72,
                        color: COLORS.ink,
                        opacity: 0.3,
                      }}
                    >
                      {current.char}
                    </div>
                  </div>
                )}
                <canvas
                  ref={recallCanvasRef}
                  width={gridSize}
                  height={gridSize}
                  style={{ position: "absolute", inset: 0, touchAction: "none", cursor: "crosshair" }}
                  onPointerDown={startRecallDrawing}
                  onPointerMove={continueRecallDrawing}
                  onPointerUp={endRecallDrawing}
                  onPointerLeave={endRecallDrawing}
                />
              </div>
            ) : status === "dots" ? (
              <div style={{ width: gridSize, height: gridSize, position: "relative", border: `1px solid ${COLORS.hairline}`, borderRadius: 10 }}>
                <svg width={gridSize} height={gridSize} style={{ position: "absolute", inset: 0, opacity: 0.9 }}>
                  <rect x={0} y={0} width={gridSize} height={gridSize} fill={COLORS.card} />
                  <line x1={mid} y1={inset} x2={mid} y2={far} stroke={COLORS.inkSoft} strokeWidth="1.2" strokeDasharray="4 4" />
                  <line x1={inset} y1={mid} x2={far} y2={mid} stroke={COLORS.inkSoft} strokeWidth="1.2" strokeDasharray="4 4" />
                  <line x1={inset} y1={inset} x2={far} y2={far} stroke={COLORS.inkSoft} strokeWidth="1.2" strokeDasharray="3 5" />
                  <line x1={far} y1={inset} x2={inset} y2={far} stroke={COLORS.inkSoft} strokeWidth="1.2" strokeDasharray="3 5" />
                </svg>
                {revealOn && dotCharData && dotSvgMatrix && (
                  <svg
                    width={gridSize}
                    height={gridSize}
                    style={{ position: "absolute", inset: 0, pointerEvents: "none", opacity: 0.3 }}
                  >
                    <g transform={dotSvgMatrix}>
                      {dotCharData.strokes.map((d, i) => (
                        <path key={i} d={d} fill={COLORS.ink} />
                      ))}
                    </g>
                  </svg>
                )}
                <canvas
                  ref={dotsCanvasRef}
                  width={gridSize}
                  height={gridSize}
                  style={{ position: "absolute", inset: 0, touchAction: "none", cursor: "crosshair" }}
                  onPointerDown={startDotDrawing}
                  onPointerMove={continueDotDrawing}
                  onPointerUp={endDotDrawing}
                  onPointerLeave={endDotDrawing}
                />
                {dotStartPoint && (
                  <div
                    style={{
                      position: "absolute",
                      left: dotStartPoint.x - 8,
                      top: dotStartPoint.y - 8,
                      width: 16,
                      height: 16,
                      borderRadius: "50%",
                      background: "#2563EB",
                      border: "2px solid white",
                      boxShadow: "0 1px 3px rgba(0,0,0,0.3)",
                      pointerEvents: "none",
                    }}
                  />
                )}
                {dotEndPoint && (
                  <div
                    style={{
                      position: "absolute",
                      left: dotEndPoint.x - 8,
                      top: dotEndPoint.y - 8,
                      width: 16,
                      height: 16,
                      borderRadius: "50%",
                      background: "#C0392B",
                      border: "2px solid white",
                      boxShadow: "0 1px 3px rgba(0,0,0,0.3)",
                      pointerEvents: "none",
                    }}
                  />
                )}
                {!dotCharData && (
                  <div
                    style={{
                      position: "absolute",
                      inset: 0,
                      display: "flex",
                      alignItems: "center",
                      justifyContent: "center",
                      fontSize: 12.5,
                      color: COLORS.inkSoft,
                    }}
                  >
                    {t("loading", meaningDisplay)}
                  </div>
                )}
              </div>
            ) : (
              <div style={{ width: gridSize, height: gridSize, position: "relative", border: `1px solid ${COLORS.hairline}`, borderRadius: 10 }}>
                <svg width={gridSize} height={gridSize} style={{ position: "absolute", inset: 0, opacity: 0.9 }}>
                  <rect x={0} y={0} width={gridSize} height={gridSize} fill={COLORS.card} />
                  <line x1={mid} y1={inset} x2={mid} y2={far} stroke={COLORS.inkSoft} strokeWidth="1.2" strokeDasharray="4 4" />
                  <line x1={inset} y1={mid} x2={far} y2={mid} stroke={COLORS.inkSoft} strokeWidth="1.2" strokeDasharray="4 4" />
                  <line x1={inset} y1={inset} x2={far} y2={far} stroke={COLORS.inkSoft} strokeWidth="1.2" strokeDasharray="3 5" />
                  <line x1={far} y1={inset} x2={inset} y2={far} stroke={COLORS.inkSoft} strokeWidth="1.2" strokeDasharray="3 5" />
                </svg>
                <div ref={targetRef} style={{ position: "absolute", inset: 0, touchAction: "none" }} />
              </div>
            )}
          </div>

          {status === "dots" && dotCharData && (
            <div style={{ textAlign: "center", marginBottom: 12 }}>
              {completedDotStrokes.length >= dotTotalStrokes ? (
                <div style={{ color: COLORS.bamboo, fontWeight: 700, fontSize: 14 }}>
                  {t("wp_dots_all_done", meaningDisplay, dotTotalStrokes)}
                </div>
              ) : (
                <>
                  <div style={{ fontSize: 12.5, color: COLORS.inkSoft, marginBottom: 4 }}>
                    {t("wp_dots_stroke_progress", meaningDisplay, dotStrokeIndex + 1, dotTotalStrokes)} {t("wp_dots_connect_prefix", meaningDisplay)}{" "}
                    <span style={{ color: "#2563EB", fontWeight: 700 }}>{t("wp_dots_green", meaningDisplay)}</span> {t("wp_dots_start_label", meaningDisplay)}{" "}
                    {t("wp_dots_to", meaningDisplay)}{" "}
                    <span style={{ color: "#C0392B", fontWeight: 700 }}>{t("wp_dots_red", meaningDisplay)}</span> {t("wp_dots_end_label", meaningDisplay)}
                  </div>
                  {dotFeedback === "correct" && (
                    <div style={{ color: COLORS.bamboo, fontWeight: 700, fontSize: 13.5 }}>{t("wp_dots_correct", meaningDisplay)}</div>
                  )}
                  {dotFeedback === "wrong" && (
                    <div style={{ color: COLORS.error, fontWeight: 700, fontSize: 13.5 }}>{t("wp_dots_wrong", meaningDisplay)}</div>
                  )}
                </>
              )}
            </div>
          )}

          {status === "dots" && dotCharData && (
            <div style={{ display: "flex", justifyContent: "center", flexWrap: "wrap", gap: 8, marginBottom: 16 }}>
              <button
                type="button"
                onClick={() => setRevealOn((r) => !r)}
                className={revealOn ? "seal-btn" : "ghost-btn"}
                style={{ ...(revealOn ? sealBtnStyle : ghostBtnStyle), padding: "8px 14px", fontSize: 12.5 }}
              >
                {t("wp_reveal", meaningDisplay)}
              </button>
              <button
                type="button"
                onClick={undoLastDotStroke}
                disabled={completedDotStrokes.length === 0}
                className="ghost-btn"
                style={{ ...ghostBtnStyle, padding: "8px 14px", fontSize: 12.5, opacity: completedDotStrokes.length === 0 ? 0.4 : 1 }}
              >
                {t("wp_undo_stroke", meaningDisplay)}
              </button>
            </div>
          )}

          {status === "dots" && !dotCharData && dotFeedback === "load-error" && (
            <div style={{ textAlign: "center", color: COLORS.inkSoft, fontSize: 13, marginBottom: 12 }}>
              {t("wp_dots_load_error", meaningDisplay, current.char)}
            </div>
          )}

          <div style={{ display: "flex", justifyContent: "center", alignItems: "center", gap: 8, marginBottom: 16 }}>
            <span style={{ fontSize: 11, color: COLORS.inkSoft }}>{t("wp_brush_size", meaningDisplay)}</span>
            {[
              { id: "thin", label: t("wp_brush_thin", meaningDisplay) },
              { id: "normal", label: t("wp_brush_normal", meaningDisplay) },
              { id: "thick", label: t("wp_brush_thick", meaningDisplay) },
            ].map((opt) => (
              <button
                key={opt.id}
                type="button"
                onClick={() => setBrushSize(opt.id)}
                style={{
                  fontSize: 11.5,
                  fontWeight: 700,
                  padding: "4px 10px",
                  borderRadius: 999,
                  border: `1px solid ${brushSize === opt.id ? COLORS.seal : COLORS.grid}`,
                  background: brushSize === opt.id ? "rgba(49,112,142,0.08)" : "transparent",
                  color: brushSize === opt.id ? COLORS.seal : COLORS.inkSoft,
                  cursor: "pointer",
                }}
              >
                {opt.label}
              </button>
            ))}
          </div>

          {status === "error" && (
            <div style={{ textAlign: "center", color: COLORS.inkSoft, fontSize: 13, marginBottom: 16 }}>
              {t("wp_no_stroke_data", meaningDisplay, current.char)}
            </div>
          )}

          {status === "done-char" && (
            <div style={{ textAlign: "center", color: COLORS.bamboo, fontSize: 14, fontWeight: 700, marginBottom: 16 }}>
              {mistakes === 0 ? t("wp_complete_no_mistakes", meaningDisplay) : t("wp_complete_with_mistakes", meaningDisplay, mistakes)}
            </div>
          )}

          {status === "recall" && (
            <div style={{ display: "flex", justifyContent: "center", flexWrap: "wrap", gap: 8, marginBottom: 16 }}>
              <button
                type="button"
                onClick={() => setRevealOn((r) => !r)}
                className={revealOn ? "seal-btn" : "ghost-btn"}
                style={{ ...(revealOn ? sealBtnStyle : ghostBtnStyle), padding: "8px 14px", fontSize: 12.5 }}
              >
                {t("wp_reveal", meaningDisplay)}
              </button>
              <button type="button" onClick={undoLastStroke} className="ghost-btn" style={{ ...ghostBtnStyle, padding: "8px 14px", fontSize: 12.5 }}>
                {t("wp_recall_erase", meaningDisplay)}
              </button>
              <button type="button" onClick={clearRecallCanvas} className="ghost-btn" style={{ ...ghostBtnStyle, padding: "8px 14px", fontSize: 12.5 }}>
                {t("wp_clear_all", meaningDisplay)}
              </button>
            </div>
          )}

          {status === "demo" ? (
            <div style={{ display: "flex", justifyContent: "center", gap: 8, marginBottom: 16 }}>
              <button type="button" onClick={replayDemo} className="ghost-btn" style={{ ...ghostBtnStyle, padding: "8px 16px", fontSize: 12.5 }}>
                {t("wp_replay", meaningDisplay)}
              </button>
              <button type="button" onClick={beginWriting} className="seal-btn" style={{ ...sealBtnStyle, padding: "8px 16px", fontSize: 12.5 }}>
                {t("wp_begin_writing", meaningDisplay)}
              </button>
            </div>
          ) : status === "dots" ? (
            <div style={{ display: "flex", justifyContent: "center", gap: 8, marginBottom: 16 }}>
              {completedDotStrokes.length >= dotTotalStrokes && dotTotalStrokes > 0 ? (
                <button type="button" onClick={enterRecallMode} className="seal-btn" style={{ ...sealBtnStyle, padding: "8px 16px", fontSize: 12.5 }}>
                  {t("wp_write_from_memory", meaningDisplay)}
                </button>
              ) : (
                <button type="button" onClick={enterRecallMode} className="ghost-btn" style={{ ...ghostBtnStyle, padding: "8px 16px", fontSize: 12.5 }}>
                  {t("wp_skip_step", meaningDisplay)}
                </button>
              )}
            </div>
          ) : status === "recall" ? (
            <div style={{ display: "flex", justifyContent: "center", gap: 8, marginBottom: 16 }}>
              <button type="button" onClick={endSession} className="ghost-btn" style={{ ...ghostBtnStyle, padding: "8px 16px", fontSize: 12.5 }}>
                {t("wp_choose_another", meaningDisplay)}
              </button>
              <button type="button" onClick={nextChar} className="seal-btn" style={{ ...sealBtnStyle, padding: "8px 16px", fontSize: 12.5 }}>
                {t("wp_next_char", meaningDisplay)}
              </button>
            </div>
          ) : (
            <div style={{ display: "flex", justifyContent: "center", gap: 8, marginBottom: 16 }}>
              <button type="button" onClick={retryChar} className="ghost-btn" style={{ ...ghostBtnStyle, padding: "8px 16px", fontSize: 12.5 }}>
                {t("wp_retrace", meaningDisplay)}
              </button>
              <button type="button" onClick={enterDotConnectMode} className="seal-btn" style={{ ...sealBtnStyle, padding: "8px 16px", fontSize: 12.5 }}>
                {t("wp_connect_dots", meaningDisplay)}
              </button>
            </div>
          )}

          <div style={{ textAlign: "center" }}>
            <button type="button" onClick={endSession} className="ghost-btn" style={{ ...ghostBtnStyle, padding: "6px 16px", fontSize: 12 }}>
              {t("wp_end", meaningDisplay)}
            </button>
          </div>
        </div>
      ) : null}
    </div>
  );
}


/* ================= ADD TAB ================= */
function AddTab({
  bushouList,
  characterList,
  wordList,
  customWords,
  onAddCharacter,
  onAddBushou,
  onUpdateCharacter,
  onDeleteCharacter,
  onAddWord,
  onDeleteWord,
  userId,
  onRequireAuth,
  onViewPremium,
  onQuotaUpdate,
  meaningDisplay,
}) {
  const [charInput, setCharInput] = useState("");
  const [meaning, setMeaning] = useState("");
  const [meaningVi, setMeaningVi] = useState("");
  const [pinyin, setPinyin] = useState("");
  const [sv, setSv] = useState("");
  const [wantMeaningEn, setWantMeaningEn] = useState(true);
  const [wantMeaningVi, setWantMeaningVi] = useState(true);
  const [wantSv, setWantSv] = useState(true);
  const [showAuthModal, setShowAuthModal] = useState(false);
  const [showNote, setShowNote] = useState(false);
  const [limitInfo, setLimitInfo] = useState(null); // { count, limit, tier } | null
  const [selectedLists, setSelectedLists] = useState([]); // a word can belong to more than one list
  const [listTypeahead, setListTypeahead] = useState("");
  const [message, setMessage] = useState(null);
  const [lookupStatus, setLookupStatus] = useState("idle"); // idle | loading | error
  const [components, setComponents] = useState([]); // chars making up charInput — from auto-lookup and/or manual entry
  const [compInput, setCompInput] = useState("");
  const [newCompDraft, setNewCompDraft] = useState(null); // { char } awaiting manual details
  const [ncPinyin, setNcPinyin] = useState("");
  const [ncMeaning, setNcMeaning] = useState("");
  const [ncSv, setNcSv] = useState("");
  const [ncStrokes, setNcStrokes] = useState("");
  const lastLookedUpRef = useRef("");

  const existingLists = useMemo(() => {
    const set = new Set();
    characterList.forEach((c) => getLists(c).forEach((l) => set.add(l.trim())));
    return Array.from(set).sort((a, b) => a.localeCompare(b, "vi"));
  }, [characterList]);

  function addList(name) {
    const trimmed = name.trim();
    if (!trimmed) return;
    setSelectedLists((prev) => (prev.includes(trimmed) ? prev : [...prev, trimmed]));
    setListTypeahead("");
  }

  function toggleList(name) {
    setSelectedLists((prev) => (prev.includes(name) ? prev.filter((l) => l !== name) : [...prev, name]));
  }

  function removeList(name) {
    setSelectedLists((prev) => prev.filter((l) => l !== name));
  }

  function resetForm() {
    setCharInput("");
    setMeaning("");
    setMeaningVi("");
    setPinyin("");
    setSv("");
    setComponents([]);
    setCompInput("");
    setNewCompDraft(null);
    lastLookedUpRef.current = "";
    // keep selectedLists so consecutive entries stay in the same list(s)
  }

  function addComponentManually() {
    const ch = compInput.trim();
    if (!ch) return;
    const known = bushouList.find((b) => b.char === ch);
    if (known) {
      setComponents((prev) => [...prev, ch]);
      setCompInput("");
    } else {
      setNewCompDraft({ char: ch });
      setNcPinyin("");
      setNcMeaning("");
      setNcSv("");
      setNcStrokes("");
    }
  }

  function confirmNewComponent() {
    if (!newCompDraft) return;
    if (!ncPinyin.trim() || !ncMeaning.trim() || !ncSv.trim()) {
      setMessage({ type: "error", text: "Điền đủ pinyin, nghĩa, và âm Hán Việt cho bộ thủ mới." });
      return;
    }
    const strokesNum = parseInt(ncStrokes, 10);
    if (onAddBushou) {
      onAddBushou({
        char: newCompDraft.char,
        pinyin: ncPinyin.trim(),
        meaning: ncMeaning.trim(),
        sv: ncSv.trim(),
        strokes: Number.isFinite(strokesNum) && strokesNum > 0 ? strokesNum : undefined,
      });
    }
    setComponents((prev) => [...prev, newCompDraft.char]);
    setNewCompDraft(null);
    setCompInput("");
    setNcStrokes("");
    setMessage(null);
  }

  function removeComponent(idx) {
    setComponents((prev) => prev.filter((_, i) => i !== idx));
  }

  async function autoFill(char, { overwrite } = { overwrite: false }) {
    const target = char.trim();
    if (!target) return;
    if (!userId) {
      setShowAuthModal(true);
      return;
    }
    // If this character already exists with full data, use it directly
    // instead of spending a paid lookup on data we already have.
    const existing = characterList.find((c) => c.char === target);
    if (existing && existing.pinyin && existing.meaning && existing.sv) {
      lastLookedUpRef.current = target;
      if (existing.pinyin && (overwrite || !pinyin.trim())) setPinyin(existing.pinyin);
      if (wantMeaningEn && existing.meaning && (overwrite || !meaning.trim())) setMeaning(existing.meaning);
      if (wantMeaningVi && existing.meaning_vi && (overwrite || !meaningVi.trim())) setMeaningVi(existing.meaning_vi);
      if (wantSv && existing.sv && (overwrite || !sv.trim())) setSv(existing.sv);
      if (Array.isArray(existing.components) && existing.components.length > 0 && (overwrite || components.length === 0)) {
        setComponents(existing.components);
      }
      return;
    }
    setLookupStatus("loading");
    try {
      const authHeaders = await getAuthHeaders();
      if (!authHeaders) {
        setLookupStatus("idle");
        setShowAuthModal(true);
        return;
      }
      const response = await fetch("/.netlify/functions/lookup-character", {
        method: "POST",
        headers: { "Content-Type": "application/json", ...authHeaders },
        body: JSON.stringify({ char: target }),
      });
      if (!response.ok) {
        const errBody = await response.json().catch(() => ({}));
        if (response.status === 401) {
          setLookupStatus("idle");
          setShowAuthModal(true);
          return;
        }
        if (response.status === 403 && (errBody.error === "LIMIT_REACHED" || errBody.error === "DISABLED")) {
          setLookupStatus("idle");
          setLimitInfo({ count: errBody.lookup_count, limit: errBody.lookup_limit, reason: errBody.error });
          onQuotaUpdate && onQuotaUpdate(errBody.lookup_count, errBody.lookup_limit);
          return;
        }
        throw new Error(errBody.error || `Lookup failed (${response.status})`);
      }
      const data = await response.json();
      if (typeof data.lookup_count === "number") {
        onQuotaUpdate && onQuotaUpdate(data.lookup_count, data.lookup_limit);
      }
      const text = (data.content || []).map((b) => b.text || "").join("");
      const clean = text.replace(/```json|```/g, "").trim();
      const parsed = JSON.parse(clean);

      if (parsed.pinyin && (overwrite || !pinyin.trim())) setPinyin(parsed.pinyin);
      if (wantMeaningEn && parsed.meaning && (overwrite || !meaning.trim())) setMeaning(parsed.meaning);
      if (wantMeaningVi && parsed.meaning_vi && (overwrite || !meaningVi.trim())) setMeaningVi(parsed.meaning_vi);
      if (wantSv && parsed.sino_vietnamese && (overwrite || !sv.trim())) setSv(parsed.sino_vietnamese);

      if (Array.isArray(parsed.components) && (overwrite || components.length === 0)) {
        const compChars = [];
        parsed.components.forEach((comp) => {
          if (!comp || !comp.char) return;
          compChars.push(comp.char);
          const alreadyKnown = bushouList.some((b) => b.char === comp.char);
          if (!alreadyKnown && comp.pinyin && comp.meaning && comp.sino_vietnamese && onAddBushou) {
            onAddBushou({ char: comp.char, pinyin: comp.pinyin, meaning: comp.meaning, sv: comp.sino_vietnamese });
          }
        });
        setComponents(compChars);
      }

      lastLookedUpRef.current = target;
      setLookupStatus("idle");
      if (!parsed.pinyin) {
        setMessage({ type: "error", text: `Không tra được thông tin cho "${target}". Vui lòng nhập tay.` });
      }
    } catch (err) {
      console.error("Auto-fill lookup failed:", err);
      setLookupStatus("error");
      setMessage({ type: "error", text: t("add_autofill_failed", meaningDisplay) });
    }
  }

  function handleSubmit(e) {
    e.preventDefault();
    setMessage(null);
    try {
      if (!charInput.trim() || !pinyin.trim()) {
        setMessage({ type: "error", text: t("add_fill_required", meaningDisplay) });
        return;
      }
      if (wantMeaningEn && !meaning.trim()) {
        setMessage({ type: "error", text: t("add_need_meaning_en", meaningDisplay) });
        return;
      }
      if (wantMeaningVi && !meaningVi.trim()) {
        setMessage({ type: "error", text: t("add_need_meaning_vi", meaningDisplay) });
        return;
      }
      if (wantSv && !sv.trim()) {
        setMessage({ type: "error", text: t("add_need_sv", meaningDisplay) });
        return;
      }
      if (selectedLists.length === 0) {
        setMessage({ type: "error", text: t("add_need_list", meaningDisplay) });
        return;
      }
      if (characterList.some((c) => c.char === charInput.trim())) {
        setMessage({ type: "error", text: t("add_char_exists", meaningDisplay, charInput.trim()) });
        return;
      }
      const listsToSave = selectedLists;
      const trimmedChar = charInput.trim();
      onAddCharacter({
        char: trimmedChar,
        meaning: wantMeaningEn ? meaning.trim() : "",
        meaning_vi: wantMeaningVi ? meaningVi.trim() : "",
        pinyin: pinyin.trim(),
        sv: wantSv ? sv.trim() : "",
        components: components,
        lists: listsToSave,
      });
      const playableNote = components.length > 0 ? "" : t("add_no_components_note", meaningDisplay);
      setMessage({ type: "success", text: t("add_char_success", meaningDisplay, trimmedChar, listsToSave.join(", "), playableNote) });
      resetForm();
    } catch (err) {
      console.error("Add character failed:", err);
      setMessage({
        type: "error",
        text: t("add_char_error", meaningDisplay, err && err.message ? err.message : t("add_unknown_error", meaningDisplay)),
      });
    }
  }

  return (
    <div>
      {showAuthModal && <AuthRequiredModal onClose={() => setShowAuthModal(false)} onSignIn={onRequireAuth} />}
      {limitInfo && (
        <LimitReachedModal
          onClose={() => setLimitInfo(null)}
          count={limitInfo.count}
          limit={limitInfo.limit}
          tier={limitInfo.tier || "Free"}
          reason={limitInfo.reason}
          onViewPremium={onViewPremium}
        />
      )}

      <div style={{ textAlign: "center", marginBottom: 20 }}>
        <button
          type="button"
          onClick={() => setShowNote((s) => !s)}
          className={showNote ? "seal-btn" : "ghost-btn"}
          style={{ ...(showNote ? sealBtnStyle : ghostBtnStyle), padding: "8px 18px", fontSize: 13 }}
        >
          {t("add_note_button", meaningDisplay)}
        </button>
        {showNote && (
          <div
            style={{
              background: COLORS.card,
              border: `1px solid ${COLORS.hairline}`,
              borderRadius: 14,
              padding: "16px 18px",
              marginTop: 12,
              fontSize: 13,
              color: COLORS.inkSoft,
              lineHeight: 1.7,
              textAlign: "left",
            }}
          >
            <p style={{ margin: 0, marginBottom: 10 }}>
              {t("add_note_p1", meaningDisplay)}
            </p>
            <p style={{ margin: 0, marginBottom: 10 }}>
              {t("add_note_p2", meaningDisplay)}
            </p>
            <p style={{ margin: 0, marginBottom: 10 }}>
              {t("add_note_p3", meaningDisplay)}
            </p>
            <p style={{ margin: 0, marginBottom: 10 }}>
              {t("add_note_p4", meaningDisplay)}
            </p>
            <p style={{ margin: 0 }}>
              <a href="https://hvdic.thivien.net/" target="_blank" rel="noopener noreferrer" style={{ color: COLORS.sealDark }}>
                hvdic.thivien.net
              </a>
              {" · "}
              <a href="https://zdic.net/" target="_blank" rel="noopener noreferrer" style={{ color: COLORS.sealDark }}>
                zdic.net
              </a>
            </p>
          </div>
        )}
      </div>

      <div style={{ fontSize: 13, color: COLORS.inkSoft, marginBottom: 18, textAlign: "center" }}>
        {t("add_intro", meaningDisplay)}
      </div>

      <div style={formCardStyle}>
        <FieldRow label={t("add_char_label", meaningDisplay)}>
          <input
            value={charInput}
            onChange={(e) => {
              setCharInput(e.target.value);
              if (e.target.value.trim() !== lastLookedUpRef.current) setComponents([]);
            }}
            onKeyDown={(e) => e.key === "Enter" && handleSubmit(e)}
            onBlur={() => {
              const trimmed = charInput.trim();
              if (trimmed && trimmed !== lastLookedUpRef.current && lookupStatus !== "loading") {
                autoFill(trimmed, { overwrite: false });
              }
            }}
            placeholder="例：好"
            style={{ ...inputStyle, fontFamily: "'Noto Serif SC', 'STKaiti', 'Kaiti SC', serif", fontSize: 22, width: 90, textAlign: "center" }}
            maxLength={4}
          />
          <button
            type="button"
            onClick={() => charInput.trim() && autoFill(charInput.trim(), { overwrite: true })}
            disabled={!charInput.trim() || lookupStatus === "loading"}
            className="ghost-btn"
            style={{ ...ghostBtnStyle, padding: "8px 12px", fontSize: 12.5, opacity: !charInput.trim() ? 0.4 : 1 }}
          >
            {lookupStatus === "loading" ? t("add_looking_up", meaningDisplay) : t("add_autofill", meaningDisplay)}
          </button>
        </FieldRow>
        <div style={{ fontSize: 11, color: COLORS.inkSoft, marginTop: -6, marginBottom: 14, paddingLeft: 204 }} className="autofill-hint">
          {t("add_char_field_description", meaningDisplay)}
        </div>

        <div style={{ borderTop: `1px dashed ${COLORS.grid}`, marginTop: 6, paddingTop: 14 }}>
          <div style={{ fontSize: 12.5, fontWeight: 600, color: COLORS.gold, marginBottom: 8, textTransform: "uppercase", letterSpacing: 0.8 }}>
            {t("add_components_label", meaningDisplay)}
          </div>
          <div style={{ fontSize: 11, color: COLORS.inkSoft, marginBottom: 10 }}>
            {t("add_components_description", meaningDisplay)}
          </div>

          {components.length > 0 && (
            <div style={{ display: "flex", flexWrap: "wrap", gap: 8, marginBottom: 10 }}>
              {components.map((ch, i) => (
                <div key={i} style={{ display: "flex", alignItems: "center", gap: 4 }}>
                  <Chip info={bushouList.find((b) => b.char === ch) || { char: ch }} meaningDisplay={meaningDisplay} />
                  <button type="button" onClick={() => removeComponent(i)} style={smallXStyle}>
                    ✕
                  </button>
                </div>
              ))}
            </div>
          )}

          <div style={{ display: "flex", gap: 8 }}>
            <input
              value={compInput}
              onChange={(e) => setCompInput(e.target.value)}
              placeholder={t("add_component_input_placeholder", meaningDisplay)}
              style={{ ...inputStyle, fontFamily: "'Noto Serif SC', 'STKaiti', 'Kaiti SC', serif", width: 140 }}
              onKeyDown={(e) => {
                if (e.key === "Enter") {
                  e.preventDefault();
                  addComponentManually();
                }
              }}
            />
            <button type="button" onClick={addComponentManually} className="ghost-btn" style={ghostBtnStyle}>
              {t("add_component_add_button", meaningDisplay)}
            </button>
          </div>

          {newCompDraft && (
            <div style={{ marginTop: 12, padding: 12, background: "rgba(80,133,165,0.08)", borderRadius: 8, border: `1px dashed ${COLORS.gold}` }}>
              <div style={{ fontSize: 12.5, marginBottom: 8 }}>
                {t("add_new_component_before", meaningDisplay)}{" "}
                <strong style={{ fontFamily: "'Noto Serif SC', 'STKaiti', 'Kaiti SC', serif", fontSize: 18 }}>{newCompDraft.char}</strong>{" "}
                {t("add_new_component_after", meaningDisplay)}
              </div>
              <div style={{ display: "flex", gap: 8, flexWrap: "wrap" }}>
                <input value={ncPinyin} onChange={(e) => setNcPinyin(e.target.value)} placeholder={t("add_new_comp_pinyin", meaningDisplay)} style={{ ...inputStyle, width: 100 }} />
                <input value={ncMeaning} onChange={(e) => setNcMeaning(e.target.value)} placeholder={t("add_new_comp_meaning", meaningDisplay)} style={{ ...inputStyle, width: 140 }} />
                <input value={ncSv} onChange={(e) => setNcSv(e.target.value)} placeholder={t("add_new_comp_sv", meaningDisplay)} style={{ ...inputStyle, width: 120 }} />
                <input
                  value={ncStrokes}
                  onChange={(e) => setNcStrokes(e.target.value.replace(/[^0-9]/g, ""))}
                  placeholder={t("add_new_comp_strokes", meaningDisplay)}
                  style={{ ...inputStyle, width: 130 }}
                />
                <button type="button" onClick={confirmNewComponent} className="seal-btn" style={{ ...sealBtnStyle, padding: "8px 14px", fontSize: 13 }}>
                  {t("add_confirm", meaningDisplay)}
                </button>
              </div>
            </div>
          )}
        </div>

        <FieldRow label={t("add_meaning_en_label", meaningDisplay)}>
          <div style={{ display: "flex", alignItems: "center", gap: 8, flex: 1 }}>
            <input
              value={meaning}
              onChange={(e) => setMeaning(e.target.value)}
              onKeyDown={(e) => e.key === "Enter" && handleSubmit(e)}
              placeholder="good, well"
              disabled={!wantMeaningEn}
              style={{ ...inputStyle, flex: 1, opacity: wantMeaningEn ? 1 : 0.45 }}
            />
            <label style={{ display: "flex", alignItems: "center", gap: 4, fontSize: 11, color: COLORS.inkSoft, whiteSpace: "nowrap", cursor: "pointer" }}>
              <input
                type="checkbox"
                checked={wantMeaningEn}
                onChange={(e) => {
                  setWantMeaningEn(e.target.checked);
                  if (!e.target.checked) setMeaning("");
                }}
              />
              {t("add_field_required_checkbox", meaningDisplay)}
            </label>
          </div>
        </FieldRow>

        <FieldRow label={t("add_meaning_vi_label", meaningDisplay)}>
          <div style={{ display: "flex", alignItems: "center", gap: 8, flex: 1 }}>
            <input
              value={meaningVi}
              onChange={(e) => setMeaningVi(e.target.value)}
              onKeyDown={(e) => e.key === "Enter" && handleSubmit(e)}
              placeholder="tốt"
              disabled={!wantMeaningVi}
              style={{ ...inputStyle, flex: 1, opacity: wantMeaningVi ? 1 : 0.45 }}
            />
            <label style={{ display: "flex", alignItems: "center", gap: 4, fontSize: 11, color: COLORS.inkSoft, whiteSpace: "nowrap", cursor: "pointer" }}>
              <input
                type="checkbox"
                checked={wantMeaningVi}
                onChange={(e) => {
                  setWantMeaningVi(e.target.checked);
                  if (!e.target.checked) setMeaningVi("");
                }}
              />
              {t("add_field_required_checkbox", meaningDisplay)}
            </label>
          </div>
        </FieldRow>

        <FieldRow label={t("add_pinyin_label", meaningDisplay)}>
          <input
            value={pinyin}
            onChange={(e) => setPinyin(e.target.value)}
            onKeyDown={(e) => e.key === "Enter" && handleSubmit(e)}
            placeholder="hǎo"
            style={inputStyle}
          />
        </FieldRow>

        <FieldRow label={t("add_hanviet_label", meaningDisplay)}>
          <div style={{ display: "flex", alignItems: "center", gap: 8, flex: 1 }}>
            <input
              value={sv}
              onChange={(e) => setSv(e.target.value)}
              onKeyDown={(e) => e.key === "Enter" && handleSubmit(e)}
              placeholder="hảo"
              disabled={!wantSv}
              style={{ ...inputStyle, flex: 1, opacity: wantSv ? 1 : 0.45 }}
            />
            <label style={{ display: "flex", alignItems: "center", gap: 4, fontSize: 11, color: COLORS.inkSoft, whiteSpace: "nowrap", cursor: "pointer" }}>
              <input
                type="checkbox"
                checked={wantSv}
                onChange={(e) => {
                  setWantSv(e.target.checked);
                  if (!e.target.checked) setSv("");
                }}
              />
              {t("add_field_required_checkbox", meaningDisplay)}
            </label>
          </div>
        </FieldRow>

        <FieldRow label={t("add_lists_label", meaningDisplay)}>
          <div style={{ flex: 1 }}>
            {selectedLists.length > 0 && (
              <div style={{ display: "flex", flexWrap: "wrap", gap: 6, marginBottom: 6 }}>
                {selectedLists.map((l) => (
                  <span
                    key={l}
                    style={{
                      display: "inline-flex",
                      alignItems: "center",
                      gap: 4,
                      fontSize: 12,
                      padding: "3px 6px 3px 10px",
                      borderRadius: 999,
                      border: `1px solid ${COLORS.seal}`,
                      background: "rgba(168,72,47,0.08)",
                      color: COLORS.seal,
                    }}
                  >
                    {l}
                    <button
                      type="button"
                      onClick={() => removeList(l)}
                      style={{ background: "none", border: "none", color: COLORS.seal, cursor: "pointer", fontSize: 12, lineHeight: 1, padding: 0 }}
                    >
                      ✕
                    </button>
                  </span>
                ))}
              </div>
            )}
            <div style={{ display: "flex", gap: 8 }}>
              <input
                value={listTypeahead}
                onChange={(e) => setListTypeahead(e.target.value)}
                onKeyDown={(e) => {
                  if (e.key === "Enter") {
                    e.preventDefault();
                    addList(listTypeahead);
                  }
                }}
                placeholder={t("add_list_placeholder", meaningDisplay)}
                list="existing-lists"
                style={inputStyle}
              />
              <button type="button" onClick={() => addList(listTypeahead)} className="ghost-btn" style={ghostBtnStyle}>
                {t("add_list_add_button", meaningDisplay)}
              </button>
            </div>
            <datalist id="existing-lists">
              {existingLists.map((l) => (
                <option key={l} value={l} />
              ))}
            </datalist>
          </div>
        </FieldRow>
        <div style={{ fontSize: 10.5, color: COLORS.inkSoft, marginTop: -6, marginBottom: 6, paddingLeft: 204 }} className="autofill-hint">
          {t("add_lists_description", meaningDisplay)}
        </div>

        {existingLists.length > 0 && (
          <div className="list-pills-row" style={{ display: "flex", flexWrap: "wrap", gap: 6, marginTop: -2, marginBottom: 6 }}>
            {existingLists.map((l) => (
              <button
                key={l}
                type="button"
                onClick={() => toggleList(l)}
                style={{
                  fontSize: 11.5,
                  padding: "3px 9px",
                  borderRadius: 999,
                  border: `1px solid ${selectedLists.includes(l) ? COLORS.seal : COLORS.grid}`,
                  background: selectedLists.includes(l) ? "rgba(168,72,47,0.08)" : "transparent",
                  color: selectedLists.includes(l) ? COLORS.seal : COLORS.inkSoft,
                  cursor: "pointer",
                }}
              >
                {selectedLists.includes(l) ? "✓ " : ""}
                {l}
              </button>
            ))}
          </div>
        )}

        {message && (
          <div style={{ marginTop: 14, fontSize: 13, color: message.type === "error" ? COLORS.error : COLORS.bamboo, fontWeight: 600 }}>
            {message.text}
          </div>
        )}

        <div style={{ marginTop: 18, textAlign: "center" }}>
          <button type="button" onClick={handleSubmit} className="seal-btn" style={sealBtnStyle}>
            {t("add_save", meaningDisplay)}
          </button>
        </div>
      </div>

      <div style={{ marginTop: 28 }}>
        <BulkImportPanel
          characterList={characterList}
          wordList={wordList}
          bushouList={bushouList}
          onAddCharacter={onAddCharacter}
          onAddBushou={onAddBushou}
          onUpdateCharacter={onUpdateCharacter}
          onAddWord={onAddWord}
          userId={userId}
          onRequireAuth={onRequireAuth}
          onViewPremium={onViewPremium}
          onQuotaUpdate={onQuotaUpdate}
          meaningDisplay={meaningDisplay}
        />
      </div>

      <div style={{ marginTop: 20 }}>
        <AddWordPanel
          characterList={characterList}
          wordList={wordList}
          customWords={customWords}
          bushouList={bushouList}
          onAddCharacter={onAddCharacter}
          onAddBushou={onAddBushou}
          onAddWord={onAddWord}
          onDeleteWord={onDeleteWord}
          userId={userId}
          onRequireAuth={onRequireAuth}
          onViewPremium={onViewPremium}
          onQuotaUpdate={onQuotaUpdate}
          meaningDisplay={meaningDisplay}
        />
      </div>
    </div>
  );
}

const BULK_IMPORT_MAX = 20;

/* ---------- Bulk import: paste up to 20 items, one per line, each either a
   single character or a multi-character word. Each is looked up for real
   (same lookups as the single-item auto-fill flows) and tagged with one
   list name. Existing items just get the list name appended. ---------- */
function BulkImportPanel({ characterList, wordList, bushouList, onAddCharacter, onAddBushou, onUpdateCharacter, onAddWord, userId, onRequireAuth, onViewPremium, onQuotaUpdate, meaningDisplay }) {
  const [expanded, setExpanded] = useState(false);
  const [rawInput, setRawInput] = useState("");
  const [selectedLists, setSelectedLists] = useState([]);
  const [listTypeahead, setListTypeahead] = useState("");
  const [wantMeaningEn, setWantMeaningEn] = useState(true);
  const [wantMeaningVi, setWantMeaningVi] = useState(true);
  const [wantSv, setWantSv] = useState(true);
  const [status, setStatus] = useState("idle"); // idle | running | done
  const [progress, setProgress] = useState({ done: 0, total: 0, current: "" });
  const [results, setResults] = useState([]); // [{item, kind: 'char'|'word', outcome, detail}]
  const [showAuthModal, setShowAuthModal] = useState(false);
  const [limitInfo, setLimitInfo] = useState(null); // { count, limit } | null
  const cancelRef = useRef(false);

  const existingLists = useMemo(() => {
    const set = new Set();
    characterList.forEach((c) => getLists(c).forEach((l) => set.add(l.trim())));
    wordList.forEach((w) => (w.lists || []).forEach((l) => set.add(l.trim())));
    return Array.from(set).sort((a, b) => a.localeCompare(b, "vi"));
  }, [characterList, wordList]);

  function addList(name) {
    const trimmed = name.trim();
    if (!trimmed) return;
    setSelectedLists((prev) => (prev.includes(trimmed) ? prev : [...prev, trimmed]));
    setListTypeahead("");
  }

  function toggleList(name) {
    setSelectedLists((prev) => (prev.includes(name) ? prev.filter((l) => l !== name) : [...prev, name]));
  }

  function removeList(name) {
    setSelectedLists((prev) => prev.filter((l) => l !== name));
  }

  function sleep(ms) {
    return new Promise((resolve) => setTimeout(resolve, ms));
  }

  function parseItems() {
    const lines = rawInput.split(/\r?\n/);
    const items = [];
    const seen = new Set();
    for (const rawLine of lines) {
      const cleaned = Array.from(rawLine)
        .filter((ch) => /[\u4e00-\u9fff]/.test(ch))
        .join("");
      if (cleaned && !seen.has(cleaned)) {
        seen.add(cleaned);
        items.push(cleaned);
      }
    }
    return items;
  }

  async function lookupAndAddCharacter(ch, tags, addedThisRun) {
    const authHeaders = await getAuthHeaders();
    if (!authHeaders) throw new Error("AUTH_REQUIRED");
    const response = await fetch("/.netlify/functions/lookup-character", {
      method: "POST",
      headers: { "Content-Type": "application/json", ...authHeaders },
      body: JSON.stringify({ char: ch }),
    });
    if (response.status === 401) throw new Error("AUTH_REQUIRED");
    if (response.status === 403) {
      const errBody = await response.json().catch(() => ({}));
      if (errBody.error === "LIMIT_REACHED" || errBody.error === "DISABLED") {
        const limitErr = new Error("LIMIT_REACHED");
        limitErr.lookup_count = errBody.lookup_count;
        limitErr.lookup_limit = errBody.lookup_limit;
        limitErr.reason = errBody.error;
        throw limitErr;
      }
    }
    if (!response.ok) throw new Error(`lookup failed (${response.status})`);
    const data = await response.json();
    if (typeof data.lookup_count === "number") {
      onQuotaUpdate && onQuotaUpdate(data.lookup_count, data.lookup_limit);
    }
    const text = (data.content || []).map((b) => b.text || "").join("");
    const clean = text.replace(/```json|```/g, "").trim();
    const parsed = JSON.parse(clean);
    if (!parsed.pinyin) throw new Error(t("bulk_missing_pinyin", meaningDisplay));
    if (wantMeaningEn && !parsed.meaning) throw new Error(t("bulk_missing_meaning_en", meaningDisplay));
    if (wantMeaningVi && !parsed.meaning_vi) throw new Error(t("bulk_missing_meaning_vi", meaningDisplay));
    if (wantSv && !parsed.sino_vietnamese) throw new Error(t("bulk_missing_sv", meaningDisplay));

    const compChars = [];
    (parsed.components || []).forEach((comp) => {
      if (!comp || !comp.char) return;
      compChars.push(comp.char);
      const alreadyKnown = bushouList.some((b) => b.char === comp.char) || addedThisRun.bushou.has(comp.char);
      if (!alreadyKnown && comp.pinyin && comp.meaning && comp.sino_vietnamese) {
        onAddBushou({ char: comp.char, pinyin: comp.pinyin, meaning: comp.meaning, sv: comp.sino_vietnamese });
        addedThisRun.bushou.add(comp.char);
      }
    });

    await onAddCharacter({
      char: ch,
      pinyin: parsed.pinyin || "",
      meaning: wantMeaningEn ? parsed.meaning || "" : "",
      meaning_vi: wantMeaningVi ? parsed.meaning_vi || "" : "",
      sv: wantSv ? parsed.sino_vietnamese || "" : "",
      components: compChars,
      lists: tags,
    });
    addedThisRun.chars.set(ch, compChars);
  }

  async function startImport() {
    if (!userId) {
      setShowAuthModal(true);
      return;
    }
    const items = parseItems();

    if (items.length === 0) {
      setResults([{ item: "", outcome: "error", detail: t("bulk_no_chars_found", meaningDisplay) }]);
      return;
    }
    if (items.length > BULK_IMPORT_MAX) {
      setResults([{ item: "", outcome: "error", detail: `Tối đa ${BULK_IMPORT_MAX} mục mỗi lần — bạn đã dán ${items.length}. Vui lòng bớt lại.` }]);
      return;
    }
    if (selectedLists.length === 0) {
      setResults([{ item: "", outcome: "error", detail: t("bulk_need_list", meaningDisplay) }]);
      return;
    }

    cancelRef.current = false;
    setStatus("running");
    setResults([]);
    setProgress({ done: 0, total: items.length, current: "" });

    // Tracks what THIS run has already added, so a character shared by two
    // words in the same batch (e.g. 你好 and 你们 both needing 你) isn't
    // looked up twice — characterList/bushouList props won't reflect
    // additions mid-run since this loop doesn't wait for a re-render.
    const addedThisRun = { chars: new Map(), bushou: new Set() };

    for (let i = 0; i < items.length; i++) {
      if (cancelRef.current) break;
      const item = items[i];
      const isWord = item.length >= 2;
      setProgress({ done: i, total: items.length, current: item });

      try {
        if (isWord) {
          const existingWord = wordList.find((w) => w.word === item);
          if (existingWord) {
            const existingLists = existingWord.lists || [];
            const merged = Array.from(new Set([...existingLists, ...selectedLists]));
            if (merged.length !== existingLists.length) {
              await onAddWord({ ...existingWord, lists: merged });
            }
            setResults((prev) => [...prev, { item, kind: "word", outcome: "tagged" }]);
          } else {
            const uniqueChars = Array.from(new Set(Array.from(item)));
            for (const ch of uniqueChars) {
              const existingChar = characterList.find((c) => c.char === ch);
              const ready =
                (existingChar && Array.isArray(existingChar.components) && existingChar.components.length >= 2) ||
                addedThisRun.chars.has(ch);
              if (!ready) {
                await lookupAndAddCharacter(ch, selectedLists, addedThisRun);
                await sleep(200);
              }
            }
            const wordAuthHeaders = await getAuthHeaders();
            if (!wordAuthHeaders) throw new Error("AUTH_REQUIRED");
            const wordResponse = await fetch("/.netlify/functions/lookup-word", {
              method: "POST",
              headers: { "Content-Type": "application/json", ...wordAuthHeaders },
              body: JSON.stringify({ word: item }),
            });
            if (wordResponse.status === 401) throw new Error("AUTH_REQUIRED");
            if (wordResponse.status === 403) {
              const errBody = await wordResponse.json().catch(() => ({}));
              if (errBody.error === "LIMIT_REACHED" || errBody.error === "DISABLED") {
                const limitErr = new Error("LIMIT_REACHED");
                limitErr.lookup_count = errBody.lookup_count;
                limitErr.lookup_limit = errBody.lookup_limit;
                limitErr.reason = errBody.error;
                throw limitErr;
              }
            }
            if (!wordResponse.ok) throw new Error(`word lookup failed (${wordResponse.status})`);
            const wordData = await wordResponse.json();
            if (typeof wordData.lookup_count === "number") {
              onQuotaUpdate && onQuotaUpdate(wordData.lookup_count, wordData.lookup_limit);
            }
            const wordText = (wordData.content || []).map((b) => b.text || "").join("");
            const wordClean = wordText.replace(/```json|```/g, "").trim();
            const wordParsed = JSON.parse(wordClean);
            if (!wordParsed.pinyin) throw new Error(t("bulk_missing_pinyin", meaningDisplay));
            if (wantMeaningEn && !wordParsed.meaning) throw new Error(t("bulk_missing_meaning_en", meaningDisplay));
            if (wantMeaningVi && !wordParsed.meaning_vi) throw new Error(t("bulk_missing_meaning_vi", meaningDisplay));
            if (wantSv && !wordParsed.sino_vietnamese) throw new Error(t("bulk_missing_sv", meaningDisplay));
            await onAddWord({
              word: item,
              chars: Array.from(item),
              pinyin: wordParsed.pinyin || "",
              meaning: wantMeaningEn ? wordParsed.meaning || "" : "",
              meaning_vi: wantMeaningVi ? wordParsed.meaning_vi || "" : "",
              sv: wantSv ? wordParsed.sino_vietnamese || "" : "",
              lists: selectedLists,
            });
            setResults((prev) => [...prev, { item, kind: "word", outcome: "added" }]);
          }
        } else {
          const existingChar = characterList.find((c) => c.char === item);
          if (existingChar) {
            const existingLists = getLists(existingChar);
            const merged = Array.from(new Set([...existingLists, ...selectedLists]));
            if (merged.length !== existingLists.length) {
              await onUpdateCharacter(item, { lists: merged });
            }
            setResults((prev) => [...prev, { item, kind: "char", outcome: "tagged" }]);
          } else if (addedThisRun.chars.has(item)) {
            setResults((prev) => [...prev, { item, kind: "char", outcome: "tagged" }]);
          } else {
            await lookupAndAddCharacter(item, selectedLists, addedThisRun);
            setResults((prev) => [...prev, { item, kind: "char", outcome: "added" }]);
          }
        }
      } catch (err) {
        if (err.message === "AUTH_REQUIRED") {
          console.error(`Bulk import stopped for "${item}": session expired`);
          setShowAuthModal(true);
          break;
        }
        if (err.message === "LIMIT_REACHED") {
          console.error(`Bulk import stopped for "${item}": lookup limit reached`);
          setLimitInfo({ count: err.lookup_count, limit: err.lookup_limit, reason: err.reason });
          onQuotaUpdate && onQuotaUpdate(err.lookup_count, err.lookup_limit);
          break;
        }
        console.error(`Bulk import failed for "${item}":`, err);
        setResults((prev) => [...prev, { item, outcome: "error", detail: err.message }]);
      }

      setProgress({ done: i + 1, total: items.length, current: item });
      await sleep(200);
    }

    setStatus("done");
  }

  function handleCancel() {
    cancelRef.current = true;
  }

  function handleReset() {
    setStatus("idle");
    setResults([]);
    setProgress({ done: 0, total: 0, current: "" });
  }

  const addedCount = results.filter((r) => r.outcome === "added").length;
  const taggedCount = results.filter((r) => r.outcome === "tagged").length;
  const errorResults = results.filter((r) => r.outcome === "error" && r.item);
  const parsedCount = parseItems().length;

  return (
    <div
      style={{
        background: "rgba(80,133,165,0.05)",
        border: `1px dashed ${COLORS.gold}`,
        borderRadius: 11,
        padding: "12px 14px",
        marginBottom: 18,
      }}
    >
      {showAuthModal && <AuthRequiredModal onClose={() => setShowAuthModal(false)} onSignIn={onRequireAuth} />}
      {limitInfo && (
        <LimitReachedModal
          onClose={() => setLimitInfo(null)}
          count={limitInfo.count}
          limit={limitInfo.limit}
          tier={limitInfo.tier || "Free"}
          reason={limitInfo.reason}
          onViewPremium={onViewPremium}
        />
      )}

      <button
        type="button"
        onClick={() => setExpanded((e) => !e)}
        style={{
          width: "100%",
          background: "none",
          border: "none",
          cursor: "pointer",
          fontSize: 12.5,
          fontWeight: 700,
          color: COLORS.gold,
          textTransform: "uppercase",
          letterSpacing: 0.6,
          textAlign: "center",
        }}
      >
        {expanded ? "▲" : "▼"} {t("bulk_toggle", meaningDisplay, BULK_IMPORT_MAX)}
      </button>

      {expanded && (
        <div style={{ marginTop: 12 }}>
          <div style={{ fontSize: 11.5, color: COLORS.inkSoft, marginBottom: 10 }}>
            {t("bulk_instructions", meaningDisplay, BULK_IMPORT_MAX)}
          </div>

          <textarea
            value={rawInput}
            onChange={(e) => setRawInput(e.target.value)}
            disabled={status === "running"}
            placeholder={"例：\n好\n你好\n汉语"}
            rows={6}
            style={{ ...inputStyle, width: "100%", boxSizing: "border-box", fontFamily: "'Noto Serif SC', 'STKaiti', 'Kaiti SC', serif", fontSize: 15, resize: "vertical", marginBottom: 6, whiteSpace: "pre" }}
          />
          <div style={{ fontSize: 11, color: parsedCount > BULK_IMPORT_MAX ? COLORS.error : COLORS.inkSoft, marginBottom: 12 }}>
            {t("bulk_count", meaningDisplay, parsedCount, BULK_IMPORT_MAX)}
          </div>

          <div style={{ marginBottom: 12 }}>
            <label style={{ fontSize: 11.5, color: COLORS.inkSoft, display: "block", marginBottom: 6 }}>{t("bulk_lists_label", meaningDisplay)}</label>
            {selectedLists.length > 0 && (
              <div style={{ display: "flex", flexWrap: "wrap", gap: 6, marginBottom: 6 }}>
                {selectedLists.map((l) => (
                  <span
                    key={l}
                    style={{
                      display: "inline-flex",
                      alignItems: "center",
                      gap: 4,
                      fontSize: 12,
                      padding: "3px 6px 3px 10px",
                      borderRadius: 999,
                      border: `1px solid ${COLORS.seal}`,
                      background: "rgba(49,112,142,0.08)",
                      color: COLORS.seal,
                    }}
                  >
                    {l}
                    <button
                      type="button"
                      onClick={() => removeList(l)}
                      style={{ background: "none", border: "none", color: COLORS.seal, cursor: "pointer", fontSize: 12, lineHeight: 1, padding: 0 }}
                    >
                      ✕
                    </button>
                  </span>
                ))}
              </div>
            )}
            <div style={{ display: "flex", gap: 8, marginBottom: 6 }}>
              <input
                value={listTypeahead}
                onChange={(e) => setListTypeahead(e.target.value)}
                onKeyDown={(e) => {
                  if (e.key === "Enter") {
                    e.preventDefault();
                    addList(listTypeahead);
                  }
                }}
                disabled={status === "running"}
                placeholder={t("bulk_list_placeholder", meaningDisplay)}
                list="existing-bulk-lists"
                style={inputStyle}
              />
              <button type="button" onClick={() => addList(listTypeahead)} disabled={status === "running"} className="ghost-btn" style={ghostBtnStyle}>
                {t("bulk_add", meaningDisplay)}
              </button>
            </div>
            <datalist id="existing-bulk-lists">
              {existingLists.map((l) => (
                <option key={l} value={l} />
              ))}
            </datalist>
            {existingLists.length > 0 && (
              <div style={{ display: "flex", flexWrap: "wrap", gap: 6 }}>
                {existingLists.map((l) => (
                  <button
                    key={l}
                    type="button"
                    onClick={() => toggleList(l)}
                    disabled={status === "running"}
                    style={{
                      fontSize: 11.5,
                      padding: "3px 9px",
                      borderRadius: 999,
                      border: `1px solid ${selectedLists.includes(l) ? COLORS.seal : COLORS.grid}`,
                      background: selectedLists.includes(l) ? "rgba(49,112,142,0.08)" : "transparent",
                      color: selectedLists.includes(l) ? COLORS.seal : COLORS.inkSoft,
                      cursor: "pointer",
                    }}
                  >
                    {selectedLists.includes(l) ? "✓ " : ""}
                    {l}
                  </button>
                ))}
              </div>
            )}
          </div>

          <div style={{ marginBottom: 12 }}>
            <label style={{ fontSize: 11.5, color: COLORS.inkSoft, display: "block", marginBottom: 6 }}>
              {t("bulk_required_fields_label", meaningDisplay)}
            </label>
            <div style={{ display: "flex", gap: 14, flexWrap: "wrap" }}>
              <label style={{ display: "flex", alignItems: "center", gap: 5, fontSize: 12.5, color: COLORS.ink, cursor: "pointer" }}>
                <input type="checkbox" checked={wantMeaningEn} disabled={status === "running"} onChange={(e) => setWantMeaningEn(e.target.checked)} />
                {t("bulk_field_meaning_en", meaningDisplay)}
              </label>
              <label style={{ display: "flex", alignItems: "center", gap: 5, fontSize: 12.5, color: COLORS.ink, cursor: "pointer" }}>
                <input type="checkbox" checked={wantMeaningVi} disabled={status === "running"} onChange={(e) => setWantMeaningVi(e.target.checked)} />
                {t("bulk_field_meaning_vi", meaningDisplay)}
              </label>
              <label style={{ display: "flex", alignItems: "center", gap: 5, fontSize: 12.5, color: COLORS.ink, cursor: "pointer" }}>
                <input type="checkbox" checked={wantSv} disabled={status === "running"} onChange={(e) => setWantSv(e.target.checked)} />
                {t("bulk_field_sv", meaningDisplay)}
              </label>
            </div>
          </div>

          <div style={{ display: "flex", gap: 8, marginBottom: 10, flexWrap: "wrap" }}>
            {status !== "running" ? (
              <button type="button" onClick={startImport} className="seal-btn" style={{ ...sealBtnStyle, padding: "8px 16px", fontSize: 13 }}>
                {t("bulk_start", meaningDisplay)}
              </button>
            ) : (
              <button type="button" onClick={handleCancel} className="ghost-btn" style={ghostBtnStyle}>
                {t("bulk_stop", meaningDisplay)}
              </button>
            )}
            {status === "done" && (
              <button type="button" onClick={handleReset} className="ghost-btn" style={ghostBtnStyle}>
                {t("bulk_reset", meaningDisplay)}
              </button>
            )}
          </div>

          {(status === "running" || status === "done") && progress.total > 0 && (
            <div style={{ marginBottom: 10 }}>
              <div style={{ fontSize: 12, color: COLORS.inkSoft, marginBottom: 4 }}>
                {progress.done} / {progress.total}
                {status === "running" ? ` ${t("bulk_processing", meaningDisplay, progress.current)}` : ` ${t("bulk_complete", meaningDisplay)}`}
              </div>
              <div style={{ height: 6, background: COLORS.grid, borderRadius: 3, overflow: "hidden" }}>
                <div
                  style={{
                    height: "100%",
                    width: `${(progress.done / progress.total) * 100}%`,
                    background: COLORS.seal,
                    transition: "width 0.2s ease",
                  }}
                />
              </div>
            </div>
          )}

          {results.length > 0 && results[0].item === "" && results[0].outcome === "error" && (
            <div style={{ fontSize: 12.5, fontWeight: 600, color: COLORS.error }}>{results[0].detail}</div>
          )}

          {status === "done" && !(results.length === 1 && results[0].item === "") && (
            <div style={{ fontSize: 12.5, fontWeight: 600 }}>
              <div style={{ color: COLORS.bamboo, marginBottom: 4 }}>
                {t("bulk_summary", meaningDisplay, addedCount, taggedCount)}
              </div>
              {errorResults.length > 0 && (
                <div style={{ color: COLORS.error }}>
                  {t("bulk_errors", meaningDisplay, errorResults.length, errorResults.map((r) => r.item).join(", "))}
                  <div style={{ fontWeight: 400, fontSize: 11.5, marginTop: 2 }}>
                    {t("bulk_retry_hint", meaningDisplay)}
                  </div>
                </div>
              )}
            </div>
          )}
        </div>
      )}
    </div>
  );
}

/* ---------- Rename a list across everything tagged with it — both
   characters and words. Editing a seed-sourced item creates a personal
   override the same way any other edit does, so this touches every
   character/word carrying the old name, not just custom ones. ---------- */
/* ---------- Add word: build a multi-character word from characters that
   already exist (or can be auto-filled on the spot), tag it with lists,
   and save it. Only shows/manages the current user's own custom words —
   the built-in seed words aren't editable here. ---------- */
function AddWordPanel({ characterList, wordList, customWords, bushouList, onAddCharacter, onAddBushou, onAddWord, onDeleteWord, userId, onRequireAuth, onViewPremium, onQuotaUpdate, meaningDisplay }) {
  const [expanded, setExpanded] = useState(false);
  const [wordInput, setWordInput] = useState("");
  const [pinyin, setPinyin] = useState("");
  const [meaning, setMeaning] = useState("");
  const [meaningVi, setMeaningVi] = useState("");
  const [sv, setSv] = useState("");
  const [wantMeaningEn, setWantMeaningEn] = useState(true);
  const [wantMeaningVi, setWantMeaningVi] = useState(true);
  const [wantSv, setWantSv] = useState(true);
  const [selectedLists, setSelectedLists] = useState([]);
  const [listTypeahead, setListTypeahead] = useState("");
  const [message, setMessage] = useState(null);
  const [charStatus, setCharStatus] = useState({}); // char -> "loading" | "error"
  const [wordLookupStatus, setWordLookupStatus] = useState("idle"); // idle | loading | error
  const [showAuthModal, setShowAuthModal] = useState(false);
  const [limitInfo, setLimitInfo] = useState(null); // { count, limit } | null
  const lastAutoFilledRef = useRef("");

  const chars = Array.from(wordInput).filter((ch) => /[\u4e00-\u9fff]/.test(ch));
  const uniqueChars = Array.from(new Set(chars));

  function charInfo(ch) {
    const found = characterList.find((c) => c.char === ch);
    const ready = !!(found && Array.isArray(found.components) && found.components.length >= 2);
    return { found, ready };
  }

  const allReady = chars.length >= 2 && uniqueChars.every((ch) => charInfo(ch).ready);

  const existingLists = useMemo(() => {
    const set = new Set();
    wordList.forEach((w) => (w.lists || []).forEach((l) => set.add(l.trim())));
    return Array.from(set).sort((a, b) => a.localeCompare(b, "vi"));
  }, [wordList]);

  async function handleAutoFillChar(ch) {
    if (!userId) {
      setShowAuthModal(true);
      return;
    }
    setCharStatus((prev) => ({ ...prev, [ch]: "loading" }));
    try {
      const authHeaders = await getAuthHeaders();
      if (!authHeaders) {
        setCharStatus((prev) => {
          const next = { ...prev };
          delete next[ch];
          return next;
        });
        setShowAuthModal(true);
        return;
      }
      const response = await fetch("/.netlify/functions/lookup-character", {
        method: "POST",
        headers: { "Content-Type": "application/json", ...authHeaders },
        body: JSON.stringify({ char: ch }),
      });
      if (response.status === 401) {
        setCharStatus((prev) => {
          const next = { ...prev };
          delete next[ch];
          return next;
        });
        setShowAuthModal(true);
        return;
      }
      if (response.status === 403) {
        const errBody = await response.json().catch(() => ({}));
        setCharStatus((prev) => {
          const next = { ...prev };
          delete next[ch];
          return next;
        });
        if (errBody.error === "LIMIT_REACHED" || errBody.error === "DISABLED") {
          setLimitInfo({ count: errBody.lookup_count, limit: errBody.lookup_limit, reason: errBody.error });
          onQuotaUpdate && onQuotaUpdate(errBody.lookup_count, errBody.lookup_limit);
          return;
        }
      }
      if (!response.ok) throw new Error(`lookup failed (${response.status})`);
      const data = await response.json();
      if (typeof data.lookup_count === "number") {
        onQuotaUpdate && onQuotaUpdate(data.lookup_count, data.lookup_limit);
      }
      const text = (data.content || []).map((b) => b.text || "").join("");
      const clean = text.replace(/```json|```/g, "").trim();
      const parsed = JSON.parse(clean);
      if (!parsed.pinyin && !parsed.meaning && !parsed.sino_vietnamese) throw new Error("no data returned");

      const compChars = [];
      (parsed.components || []).forEach((comp) => {
        if (!comp || !comp.char) return;
        compChars.push(comp.char);
        const alreadyKnown = bushouList.some((b) => b.char === comp.char);
        if (!alreadyKnown && comp.pinyin && comp.meaning && comp.sino_vietnamese) {
          onAddBushou({ char: comp.char, pinyin: comp.pinyin, meaning: comp.meaning, sv: comp.sino_vietnamese });
        }
      });

      await onAddCharacter({
        char: ch,
        pinyin: parsed.pinyin || "",
        meaning: parsed.meaning || "",
        meaning_vi: parsed.meaning_vi || "",
        sv: parsed.sino_vietnamese || "",
        components: compChars,
        lists: ["Chưa phân loại"],
      });
      setCharStatus((prev) => {
        const next = { ...prev };
        delete next[ch];
        return next;
      });
    } catch (err) {
      console.error(`Auto-fill failed for "${ch}":`, err);
      setCharStatus((prev) => ({ ...prev, [ch]: "error" }));
    }
  }

  async function handleAutoFillWordMeta() {
    const word = chars.join("");
    if (!word) return;
    if (!userId) {
      setShowAuthModal(true);
      return;
    }
    // If this word already exists with full data, use it directly instead
    // of spending a paid lookup on data we already have.
    const existingWord = wordList.find((w) => w.word === word);
    if (existingWord && existingWord.pinyin && existingWord.meaning && existingWord.sv) {
      if (existingWord.pinyin) setPinyin(existingWord.pinyin);
      if (wantMeaningEn && existingWord.meaning) setMeaning(existingWord.meaning);
      if (wantMeaningVi && existingWord.meaning_vi) setMeaningVi(existingWord.meaning_vi);
      if (wantSv && existingWord.sv) setSv(existingWord.sv);
      return;
    }
    setWordLookupStatus("loading");
    try {
      const authHeaders = await getAuthHeaders();
      if (!authHeaders) {
        setWordLookupStatus("idle");
        setShowAuthModal(true);
        return;
      }
      const response = await fetch("/.netlify/functions/lookup-word", {
        method: "POST",
        headers: { "Content-Type": "application/json", ...authHeaders },
        body: JSON.stringify({ word }),
      });
      if (response.status === 401) {
        setWordLookupStatus("idle");
        setShowAuthModal(true);
        return;
      }
      if (response.status === 403) {
        const errBody = await response.json().catch(() => ({}));
        setWordLookupStatus("idle");
        if (errBody.error === "LIMIT_REACHED" || errBody.error === "DISABLED") {
          setLimitInfo({ count: errBody.lookup_count, limit: errBody.lookup_limit, reason: errBody.error });
          onQuotaUpdate && onQuotaUpdate(errBody.lookup_count, errBody.lookup_limit);
          return;
        }
      }
      if (!response.ok) throw new Error(`lookup failed (${response.status})`);
      const data = await response.json();
      if (typeof data.lookup_count === "number") {
        onQuotaUpdate && onQuotaUpdate(data.lookup_count, data.lookup_limit);
      }
      const text = (data.content || []).map((b) => b.text || "").join("");
      const clean = text.replace(/```json|```/g, "").trim();
      const parsed = JSON.parse(clean);
      if (parsed.pinyin) setPinyin(parsed.pinyin);
      if (wantMeaningEn && parsed.meaning) setMeaning(parsed.meaning);
      if (wantMeaningVi && parsed.meaning_vi) setMeaningVi(parsed.meaning_vi);
      if (wantSv && parsed.sino_vietnamese) setSv(parsed.sino_vietnamese);
      if (!parsed.pinyin) {
        setMessage({ type: "error", text: `Không tra được thông tin cho từ "${word}". Vui lòng nhập tay.` });
      }
      setWordLookupStatus("idle");
    } catch (err) {
      console.error("Word lookup failed:", err);
      setWordLookupStatus("error");
      setMessage({ type: "error", text: t("word_lookup_failed", meaningDisplay) });
    }
  }

  // One button (or one blur of the word field) does everything: looks up
  // pinyin/meaning/Hán Việt for the whole word, AND auto-fills any
  // character in it that doesn't have components yet.
  async function autoFillEverything() {
    if (chars.length < 2) return;
    if (!userId) {
      setShowAuthModal(true);
      return;
    }
    setMessage(null);
    lastAutoFilledRef.current = chars.join("");
    await handleAutoFillWordMeta();
    const missing = uniqueChars.filter((ch) => !charInfo(ch).ready);
    for (const ch of missing) {
      await handleAutoFillChar(ch);
    }
  }

  function addList(name) {
    const trimmed = name.trim();
    if (!trimmed) return;
    setSelectedLists((prev) => (prev.includes(trimmed) ? prev : [...prev, trimmed]));
    setListTypeahead("");
  }

  function toggleList(name) {
    setSelectedLists((prev) => (prev.includes(name) ? prev.filter((l) => l !== name) : [...prev, name]));
  }

  function removeList(name) {
    setSelectedLists((prev) => prev.filter((l) => l !== name));
  }

  function handleSaveWord() {
    setMessage(null);
    const word = chars.join("");
    if (chars.length < 2) {
      setMessage({ type: "error", text: t("word_need_2_chars", meaningDisplay) });
      return;
    }
    if (!allReady) {
      setMessage({ type: "error", text: t("word_need_components", meaningDisplay) });
      return;
    }
    if (!pinyin.trim()) {
      setMessage({ type: "error", text: t("word_need_pinyin_meaning", meaningDisplay) });
      return;
    }
    if (wantMeaningEn && !meaning.trim()) {
      setMessage({ type: "error", text: t("add_need_meaning_en", meaningDisplay) });
      return;
    }
    if (wantMeaningVi && !meaningVi.trim()) {
      setMessage({ type: "error", text: t("add_need_meaning_vi", meaningDisplay) });
      return;
    }
    if (wantSv && !sv.trim()) {
      setMessage({ type: "error", text: t("add_need_sv", meaningDisplay) });
      return;
    }
    if (selectedLists.length === 0) {
      setMessage({ type: "error", text: t("add_need_list", meaningDisplay) });
      return;
    }
    if (wordList.some((w) => w.word === word)) {
      setMessage({ type: "error", text: t("word_exists", meaningDisplay, word) });
      return;
    }
    const listsToSave = selectedLists;
    onAddWord({
      word,
      chars,
      pinyin: pinyin.trim(),
      meaning: wantMeaningEn ? meaning.trim() : "",
      meaning_vi: wantMeaningVi ? meaningVi.trim() : "",
      sv: wantSv ? sv.trim() : "",
      lists: listsToSave,
    });
    setMessage({ type: "success", text: t("word_added_success", meaningDisplay, word) });
    setWordInput("");
    setPinyin("");
    setMeaning("");
    setMeaningVi("");
    setSv("");
  }

  return (
    <div
      style={{
        background: "rgba(80,133,165,0.05)",
        border: `1px dashed ${COLORS.gold}`,
        borderRadius: 11,
        padding: "12px 14px",
        marginBottom: 18,
      }}
    >
      {showAuthModal && <AuthRequiredModal onClose={() => setShowAuthModal(false)} onSignIn={onRequireAuth} />}
      {limitInfo && (
        <LimitReachedModal
          onClose={() => setLimitInfo(null)}
          count={limitInfo.count}
          limit={limitInfo.limit}
          tier={limitInfo.tier || "Free"}
          reason={limitInfo.reason}
          onViewPremium={onViewPremium}
        />
      )}

      <button
        type="button"
        onClick={() => setExpanded((e) => !e)}
        style={{
          width: "100%",
          background: "none",
          border: "none",
          cursor: "pointer",
          fontSize: 12.5,
          fontWeight: 700,
          color: COLORS.gold,
          textTransform: "uppercase",
          letterSpacing: 0.6,
          textAlign: "center",
        }}
      >
        {expanded ? "▲" : "▼"} {t("word_toggle", meaningDisplay)}
      </button>

      {expanded && (
        <div style={{ marginTop: 12 }}>
          <div style={{ fontSize: 11.5, color: COLORS.inkSoft, marginBottom: 10 }}>
            {t("word_instructions", meaningDisplay)}
          </div>

          <div style={{ display: "flex", gap: 8, marginBottom: 10 }}>
            <input
              value={wordInput}
              onChange={(e) => setWordInput(e.target.value)}
              onBlur={() => {
                const word = chars.join("");
                if (word.length >= 2 && word !== lastAutoFilledRef.current && wordLookupStatus !== "loading") {
                  autoFillEverything();
                }
              }}
              placeholder="例：你好"
              style={{ ...inputStyle, fontFamily: "'Noto Serif SC', 'STKaiti', 'Kaiti SC', serif", fontSize: 20 }}
            />
            <button
              type="button"
              onClick={autoFillEverything}
              disabled={chars.length < 2 || wordLookupStatus === "loading"}
              className="ghost-btn"
              style={{ ...ghostBtnStyle, padding: "8px 12px", fontSize: 12.5, opacity: chars.length < 2 ? 0.4 : 1, whiteSpace: "nowrap" }}
            >
              {wordLookupStatus === "loading" ? t("word_looking_up", meaningDisplay) : t("word_autofill", meaningDisplay)}
            </button>
          </div>

          {uniqueChars.length > 0 && (
            <div style={{ display: "flex", flexWrap: "wrap", gap: 8, marginBottom: 12 }}>
              {uniqueChars.map((ch) => {
                const { ready } = charInfo(ch);
                const status = charStatus[ch];
                return (
                  <div
                    key={ch}
                    style={{
                      display: "flex",
                      alignItems: "center",
                      gap: 6,
                      padding: "6px 10px",
                      borderRadius: 6,
                      border: `1.5px solid ${ready ? COLORS.bamboo : COLORS.error}`,
                      background: COLORS.chipBg,
                    }}
                  >
                    <span style={{ fontFamily: "'Noto Serif SC', 'STKaiti', 'Kaiti SC', serif", fontSize: 20 }}>{ch}</span>
                    {ready ? (
                      <span style={{ fontSize: 11, color: COLORS.bamboo, fontWeight: 600 }}>{t("word_ready", meaningDisplay)}</span>
                    ) : (
                      <>
                        <span style={{ fontSize: 11, color: COLORS.error }}>{t("word_no_components", meaningDisplay)}</span>
                        <button
                          type="button"
                          onClick={() => handleAutoFillChar(ch)}
                          disabled={status === "loading"}
                          className="ghost-btn"
                          style={{ ...ghostBtnStyle, padding: "3px 8px", fontSize: 10.5, borderColor: COLORS.gold, color: COLORS.gold }}
                        >
                          {status === "loading" ? t("word_looking_up", meaningDisplay) : t("word_autofill", meaningDisplay)}
                        </button>
                      </>
                    )}
                    {status === "error" && <span style={{ fontSize: 10.5, color: COLORS.error }}>{t("word_failed", meaningDisplay)}</span>}
                  </div>
                );
              })}
            </div>
          )}

          <FieldRow label={t("word_pinyin_label", meaningDisplay)}>
            <input value={pinyin} onChange={(e) => setPinyin(e.target.value)} placeholder={t("word_auto_or_manual", meaningDisplay)} style={inputStyle} />
          </FieldRow>
          <FieldRow label={t("word_meaning_en_label", meaningDisplay)}>
            <div style={{ display: "flex", alignItems: "center", gap: 8, flex: 1 }}>
              <input
                value={meaning}
                onChange={(e) => setMeaning(e.target.value)}
                placeholder={t("word_auto_or_manual", meaningDisplay)}
                disabled={!wantMeaningEn}
                style={{ ...inputStyle, flex: 1, opacity: wantMeaningEn ? 1 : 0.45 }}
              />
              <label style={{ display: "flex", alignItems: "center", gap: 4, fontSize: 11, color: COLORS.inkSoft, whiteSpace: "nowrap", cursor: "pointer" }}>
                <input
                  type="checkbox"
                  checked={wantMeaningEn}
                  onChange={(e) => {
                    setWantMeaningEn(e.target.checked);
                    if (!e.target.checked) setMeaning("");
                  }}
                />
                {t("add_field_required_checkbox", meaningDisplay)}
              </label>
            </div>
          </FieldRow>
          <FieldRow label={t("word_meaning_vi_label", meaningDisplay)}>
            <div style={{ display: "flex", alignItems: "center", gap: 8, flex: 1 }}>
              <input
                value={meaningVi}
                onChange={(e) => setMeaningVi(e.target.value)}
                placeholder={t("word_auto_or_manual", meaningDisplay)}
                disabled={!wantMeaningVi}
                style={{ ...inputStyle, flex: 1, opacity: wantMeaningVi ? 1 : 0.45 }}
              />
              <label style={{ display: "flex", alignItems: "center", gap: 4, fontSize: 11, color: COLORS.inkSoft, whiteSpace: "nowrap", cursor: "pointer" }}>
                <input
                  type="checkbox"
                  checked={wantMeaningVi}
                  onChange={(e) => {
                    setWantMeaningVi(e.target.checked);
                    if (!e.target.checked) setMeaningVi("");
                  }}
                />
                {t("add_field_required_checkbox", meaningDisplay)}
              </label>
            </div>
          </FieldRow>
          <FieldRow label={t("word_hanviet_label", meaningDisplay)}>
            <div style={{ display: "flex", alignItems: "center", gap: 8, flex: 1 }}>
              <input
                value={sv}
                onChange={(e) => setSv(e.target.value)}
                placeholder={t("word_auto_or_manual", meaningDisplay)}
                disabled={!wantSv}
                style={{ ...inputStyle, flex: 1, opacity: wantSv ? 1 : 0.45 }}
              />
              <label style={{ display: "flex", alignItems: "center", gap: 4, fontSize: 11, color: COLORS.inkSoft, whiteSpace: "nowrap", cursor: "pointer" }}>
                <input
                  type="checkbox"
                  checked={wantSv}
                  onChange={(e) => {
                    setWantSv(e.target.checked);
                    if (!e.target.checked) setSv("");
                  }}
                />
                {t("add_field_required_checkbox", meaningDisplay)}
              </label>
            </div>
          </FieldRow>

          <FieldRow label={t("word_lists_label", meaningDisplay)}>
            <div style={{ flex: 1 }}>
              {selectedLists.length > 0 && (
                <div style={{ display: "flex", flexWrap: "wrap", gap: 6, marginBottom: 6 }}>
                  {selectedLists.map((l) => (
                    <span
                      key={l}
                      style={{
                        display: "inline-flex",
                        alignItems: "center",
                        gap: 4,
                        fontSize: 12,
                        padding: "3px 6px 3px 10px",
                        borderRadius: 999,
                        border: `1px solid ${COLORS.seal}`,
                        background: "rgba(49,112,142,0.08)",
                        color: COLORS.seal,
                      }}
                    >
                      {l}
                      <button
                        type="button"
                        onClick={() => removeList(l)}
                        style={{ background: "none", border: "none", color: COLORS.seal, cursor: "pointer", fontSize: 12, lineHeight: 1, padding: 0 }}
                      >
                        ✕
                      </button>
                    </span>
                  ))}
                </div>
              )}
              <div style={{ display: "flex", gap: 8 }}>
                <input
                  value={listTypeahead}
                  onChange={(e) => setListTypeahead(e.target.value)}
                  onKeyDown={(e) => {
                    if (e.key === "Enter") {
                      e.preventDefault();
                      addList(listTypeahead);
                    }
                  }}
                  placeholder={t("word_list_placeholder", meaningDisplay)}
                  list="existing-word-lists"
                  style={inputStyle}
                />
                <button type="button" onClick={() => addList(listTypeahead)} className="ghost-btn" style={ghostBtnStyle}>
                  {t("word_add_list", meaningDisplay)}
                </button>
              </div>
              <datalist id="existing-word-lists">
                {existingLists.map((l) => (
                  <option key={l} value={l} />
                ))}
              </datalist>
              {existingLists.length > 0 && (
                <div style={{ display: "flex", flexWrap: "wrap", gap: 6, marginTop: 6 }}>
                  {existingLists.map((l) => (
                    <button
                      key={l}
                      type="button"
                      onClick={() => toggleList(l)}
                      style={{
                        fontSize: 11.5,
                        padding: "3px 9px",
                        borderRadius: 999,
                        border: `1px solid ${selectedLists.includes(l) ? COLORS.seal : COLORS.grid}`,
                        background: selectedLists.includes(l) ? "rgba(49,112,142,0.08)" : "transparent",
                        color: selectedLists.includes(l) ? COLORS.seal : COLORS.inkSoft,
                        cursor: "pointer",
                      }}
                    >
                      {selectedLists.includes(l) ? "✓ " : ""}
                      {l}
                    </button>
                  ))}
                </div>
              )}
            </div>
          </FieldRow>

          {message && (
            <div style={{ marginTop: 8, fontSize: 12.5, fontWeight: 600, color: message.type === "error" ? COLORS.error : COLORS.bamboo }}>
              {message.text}
            </div>
          )}

          <div style={{ textAlign: "center", marginTop: 12 }}>
            <button type="button" onClick={handleSaveWord} className="seal-btn" style={{ ...sealBtnStyle, padding: "8px 20px", fontSize: 13 }}>
              {t("word_save", meaningDisplay)}
            </button>
          </div>
        </div>
      )}
    </div>
  );
}

/* ---------- Searchable, filterable list of the user's own saved words —
   same search/filter pattern as CharacterListPanel, so this scales as more
   words get added instead of staying a single unsorted row. ---------- */
/* ---------- Admin-only: browse every user, adjust their tier/limit, or
   reset their usage — replaces doing the same thing via raw SQL. ---------- */
const TIER_PRESETS = { Free: 1000, Premium: 5000 };
// "Enrolled Course" isn't a fixed-limit preset like the others -- its
// limit is set manually per course/student, and it carries an extra
// course_name field the others don't use.
const ALL_TIERS = [...Object.keys(TIER_PRESETS), "Enrolled Course"];

/* ---------- Sales/info page about paid tiers and courses. Visible to
   everyone, including guests -- this is the conversion page the
   quota-exhausted and locked-list popups link to. Static content for now;
   revise the copy freely, it's just plain text/JSX below. ---------- */

function SiteFooter({ setTab, meaningDisplay }) {
  const tabLinks = [
    { id: "blog", label: t("tab_blog", meaningDisplay) },
    { id: "premium", label: t("tab_premium", meaningDisplay) },
    { id: "about", label: t("tab_about", meaningDisplay) },
    { id: "feedback", label: t("tab_feedback", meaningDisplay) },
    ...(meaningDisplay !== "vi"
      ? [
          { id: "privacy", label: t("tab_privacy", meaningDisplay) },
          { id: "terms", label: t("tab_terms", meaningDisplay) },
        ]
      : []),
  ];
  const linkStyle = {
    background: "none",
    border: "none",
    color: "#FBF9EF",
    fontSize: 14,
    fontWeight: 500,
    cursor: "pointer",
    padding: 0,
    opacity: 0.92,
    textDecoration: "none",
    fontFamily: "inherit",
  };
  return (
    <div
      style={{
        background: COLORS.sealDark,
        margin: "48px -16px -48px",
        padding: "36px 16px 30px",
      }}
    >
      <div style={{ maxWidth: 760, margin: "0 auto", textAlign: "center" }}>
        <div style={{ display: "flex", justifyContent: "center", gap: 28, flexWrap: "wrap", marginBottom: 18 }}>
          {tabLinks.map((l) => (
            <button
              key={l.id}
              type="button"
              onClick={() => {
                setTab(l.id);
                window.scrollTo({ top: 0, behavior: "smooth" });
              }}
              style={linkStyle}
            >
              {l.label}
            </button>
          ))}
        </div>
        <div style={{ fontSize: 12, color: "#FBF9EF", opacity: 0.6 }}>{t("footer_copyright", meaningDisplay)}</div>
      </div>
    </div>
  );
}

const BLOG_CATEGORY_COLORS = {
  news: { accent: "#31708E", tint: "#E8F0F4" },
  resources: { accent: "#5C7A5A", tint: "#EEF2ED" },
  founder: { accent: "#6B5C7A", tint: "#EFEDF2" },
};

function BlogComments({ postId, meaningDisplay }) {
  const [comments, setComments] = useState(null); // null = loading
  const [name, setName] = useState("");
  const [email, setEmail] = useState("");
  const [message, setMessage] = useState("");
  const [status, setStatus] = useState("idle"); // idle | sending | done | error
  const [isGuest, setIsGuest] = useState(null);

  useEffect(() => {
    loadComments();
    (async () => {
      const {
        data: { user },
      } = await supabase.auth.getUser();
      setIsGuest(!user);
      if (user) setName(user.email.split("@")[0]);
    })();
    // eslint-disable-next-line react-hooks/exhaustive-deps
  }, [postId]);

  async function loadComments() {
    const { data, error } = await supabase
      .from("blog_comments")
      .select("id, name, message, created_at")
      .eq("post_id", postId)
      .eq("status", "approved")
      .order("created_at", { ascending: true });
    if (!error) setComments(data || []);
  }

  async function handleSubmit(e) {
    e.preventDefault();
    if (!name.trim() || !message.trim()) return;
    setStatus("sending");
    try {
      const {
        data: { user },
      } = await supabase.auth.getUser();
      const { error } = await supabase.from("blog_comments").insert({
        post_id: postId,
        name: name.trim(),
        email: isGuest ? email.trim() || null : user ? user.email : null,
        message: message.trim(),
        user_id: user ? user.id : null,
      });
      if (error) throw error;
      setStatus("done");
      setMessage("");
    } catch (err) {
      console.error("Could not submit comment:", err);
      setStatus("error");
    }
  }

  return (
    <div style={{ marginTop: 24, paddingTop: 18, borderTop: `1px dashed ${COLORS.grid}` }}>
      <div style={{ fontSize: 13, fontWeight: 700, color: COLORS.ink, marginBottom: 12 }}>
        {t("blog_comments_title", meaningDisplay, comments ? comments.length : 0)}
      </div>

      {comments === null ? (
        <div style={{ fontSize: 12.5, color: COLORS.inkSoft }}>{t("loading", meaningDisplay)}</div>
      ) : comments.length === 0 ? (
        <div style={{ fontSize: 12.5, color: COLORS.inkSoft, marginBottom: 16 }}>{t("blog_comments_none", meaningDisplay)}</div>
      ) : (
        <div style={{ display: "flex", flexDirection: "column", gap: 12, marginBottom: 20 }}>
          {comments.map((c) => (
            <div key={c.id} style={{ background: COLORS.chipBg, borderRadius: 10, padding: "10px 14px" }}>
              <div style={{ display: "flex", justifyContent: "space-between", alignItems: "baseline", marginBottom: 4 }}>
                <span style={{ fontSize: 12.5, fontWeight: 700, color: COLORS.ink }}>{c.name}</span>
                <span style={{ fontSize: 10.5, color: COLORS.metadata }}>{new Date(c.created_at).toLocaleDateString()}</span>
              </div>
              <div style={{ fontSize: 12.5, color: COLORS.inkSoft, lineHeight: 1.5, whiteSpace: "pre-wrap" }}>{c.message}</div>
            </div>
          ))}
        </div>
      )}

      {status === "done" ? (
        <div style={{ fontSize: 12.5, fontWeight: 600, color: COLORS.seal }}>{t("blog_comment_thanks", meaningDisplay)}</div>
      ) : (
        <form onSubmit={handleSubmit}>
          <div style={{ fontSize: 12, fontWeight: 700, color: COLORS.ink, marginBottom: 8 }}>{t("blog_leave_comment", meaningDisplay)}</div>
          <input
            value={name}
            onChange={(e) => setName(e.target.value)}
            placeholder={t("blog_comment_name_placeholder", meaningDisplay)}
            style={{ ...inputStyle, width: "100%", boxSizing: "border-box", fontSize: 12.5, marginBottom: 6 }}
          />
          {isGuest && (
            <input
              value={email}
              onChange={(e) => setEmail(e.target.value)}
              placeholder={t("feedback_email_placeholder", meaningDisplay)}
              style={{ ...inputStyle, width: "100%", boxSizing: "border-box", fontSize: 12.5, marginBottom: 6 }}
            />
          )}
          <textarea
            value={message}
            onChange={(e) => setMessage(e.target.value)}
            placeholder={t("blog_comment_placeholder", meaningDisplay)}
            rows={3}
            style={{ ...inputStyle, width: "100%", boxSizing: "border-box", fontSize: 12.5, resize: "vertical", marginBottom: 8 }}
          />
          <div style={{ fontSize: 11, color: COLORS.metadata, marginBottom: 8 }}>{t("blog_comment_moderation_note", meaningDisplay)}</div>
          <button type="submit" className="seal-btn" style={{ ...sealBtnStyle, padding: "6px 16px", fontSize: 12.5 }} disabled={status === "sending"}>
            {status === "sending" ? t("feedback_sending", meaningDisplay) : t("blog_comment_submit", meaningDisplay)}
          </button>
          {status === "error" && <div style={{ fontSize: 11, color: COLORS.error, marginTop: 6 }}>{t("feedback_error", meaningDisplay)}</div>}
        </form>
      )}
    </div>
  );
}

function BlogTab({ meaningDisplay }) {
  const [posts, setPosts] = useState([]);
  const [loading, setLoading] = useState(true);
  const [category, setCategory] = useState("all");
  const [expandedId, setExpandedId] = useState(null);

  useEffect(() => {
    loadPosts();
    // eslint-disable-next-line react-hooks/exhaustive-deps
  }, []);

  async function loadPosts() {
    setLoading(true);
    const { data, error } = await supabase
      .from("blog_posts")
      .select("id, title, body, category, external_link, published, created_at")
      .eq("published", true)
      .order("created_at", { ascending: false });
    if (!error) setPosts(data || []);
    setLoading(false);
  }

  const categories = [
    { id: "all", label: t("blog_cat_all", meaningDisplay) },
    { id: "news", label: t("blog_cat_news", meaningDisplay) },
    { id: "resources", label: t("blog_cat_resources", meaningDisplay) },
    { id: "founder", label: t("blog_cat_founder", meaningDisplay) },
  ];
  const filtered = category === "all" ? posts : posts.filter((p) => p.category === category);

  return (
    <div style={{ maxWidth: 680, margin: "0 auto" }}>
      <div style={{ fontSize: 22, fontWeight: 700, color: COLORS.ink, marginBottom: 18, textAlign: "center" }}>
        {t("blog_page_title", meaningDisplay)}
      </div>

      <div className="side-nav-layout" style={{ display: "flex", gap: 24, alignItems: "flex-start" }}>
        <div className="side-nav-menu" style={{ display: "flex", flexDirection: "column", gap: 4, width: 170, flexShrink: 0 }}>
          {categories.map((c) => (
            <button
              key={c.id}
              type="button"
              onClick={() => {
                setCategory(c.id);
                setExpandedId(null);
              }}
              style={{
                textAlign: "left",
                background: category === c.id ? COLORS.chipBg : "none",
                border: "none",
                borderLeft: `3px solid ${category === c.id ? COLORS.seal : "transparent"}`,
                color: category === c.id ? COLORS.ink : COLORS.inkSoft,
                fontWeight: category === c.id ? 700 : 500,
                fontSize: 13,
                padding: "8px 12px",
                cursor: "pointer",
                borderRadius: 6,
              }}
            >
              {c.label}
            </button>
          ))}
        </div>

        <div style={{ flex: 1, minWidth: 0 }}>
          {loading ? (
            <div style={{ textAlign: "center", color: COLORS.inkSoft, padding: 30 }}>{t("loading", meaningDisplay)}</div>
          ) : filtered.length === 0 ? (
            <div style={{ textAlign: "center", color: COLORS.inkSoft, padding: 30 }}>{t("blog_empty", meaningDisplay)}</div>
          ) : (
            <div style={{ display: "flex", flexDirection: "column", gap: 14 }}>
              {filtered.map((post) => {
                const isExpanded = expandedId === post.id;
                const plainText = stripHtml(post.body);
                const preview = plainText.length > 140 ? plainText.slice(0, 140).trim() + "…" : plainText;
                const readMins = Math.max(1, Math.round(plainText.split(/\s+/).filter(Boolean).length / 200));
                const catColor = BLOG_CATEGORY_COLORS[post.category] || BLOG_CATEGORY_COLORS.news;
                const catLabel = t(`blog_cat_${post.category}`, meaningDisplay);
                return (
                  <div
                    key={post.id}
                    style={{
                      background: COLORS.card,
                      border: `1px solid ${COLORS.hairline}`,
                      borderLeft: `3px solid ${catColor.accent}`,
                      borderRadius: 14,
                      padding: "18px 20px",
                    }}
                  >
                    <div style={{ display: "flex", alignItems: "center", gap: 10, marginBottom: 12, flexWrap: "wrap" }}>
                      <span
                        style={{
                          fontSize: 10.5,
                          fontWeight: 700,
                          letterSpacing: 0.04,
                          textTransform: "uppercase",
                          color: catColor.accent,
                          background: catColor.tint,
                          padding: "3px 8px",
                          borderRadius: 5,
                        }}
                      >
                        {catLabel}
                      </span>
                      <span style={{ fontSize: 11, color: COLORS.metadata }}>
                        {new Date(post.created_at).toLocaleDateString()} · {t("blog_read_time", meaningDisplay, readMins)}
                      </span>
                    </div>
                    <div
                      style={{ fontFamily: "'Noto Serif', serif", fontSize: 19, fontWeight: 700, color: COLORS.ink, marginBottom: 20, lineHeight: 1.35 }}
                      dangerouslySetInnerHTML={{ __html: enhanceCjkInHtml(post.title) }}
                    />
                    {isExpanded ? (
                      <div
                        style={{ fontSize: 14, color: COLORS.inkSoft, lineHeight: 1.7 }}
                        dangerouslySetInnerHTML={{ __html: enhanceCjkInHtml(post.body) }}
                      />
                    ) : (
                      <div
                        style={{ fontSize: 14, color: COLORS.inkSoft, lineHeight: 1.7, textAlign: "justify" }}
                        dangerouslySetInnerHTML={{ __html: enhanceCjkInHtml(preview) }}
                      />
                    )}
                    {isExpanded && post.external_link && (
                      <div style={{ marginTop: 12 }}>
                        <a href={post.external_link} target="_blank" rel="noopener noreferrer" style={{ color: COLORS.seal, fontWeight: 600, fontSize: 13.5 }}>
                          {t("blog_external_link", meaningDisplay)}
                        </a>
                      </div>
                    )}
                    {plainText.length > 140 && (
                      <div style={{ textAlign: "right" }}>
                        <button
                          type="button"
                          onClick={() => setExpandedId(isExpanded ? null : post.id)}
                          style={{ background: "none", border: "none", color: COLORS.seal, fontWeight: 600, fontSize: 13, cursor: "pointer", padding: 0, marginTop: 10 }}
                        >
                          {isExpanded ? t("blog_collapse", meaningDisplay) : t("blog_read_more", meaningDisplay)}
                        </button>
                      </div>
                    )}
                    {isExpanded && <BlogComments postId={post.id} meaningDisplay={meaningDisplay} />}
                  </div>
                );
              })}
            </div>
          )}
        </div>
      </div>
    </div>
  );
}

function AboutTab({ meaningDisplay }) {
  return (
    <div style={{ padding: "20px 4px" }}>
      <div style={{ fontSize: 22, fontWeight: 700, color: COLORS.ink, marginBottom: 14, textAlign: "center" }}>
        {t("about_title", meaningDisplay)}
      </div>
      <div style={{ fontSize: 15, color: COLORS.inkSoft, lineHeight: 1.7, maxWidth: 560, margin: "0 auto", textAlign: "justify" }}>
        {t("about_body", meaningDisplay)}
      </div>
      <div style={{ fontSize: 15, color: COLORS.inkSoft, lineHeight: 1.7, maxWidth: 560, margin: "16px auto 0", textAlign: "justify" }}>
        {t("about_visit_line", meaningDisplay)}{" "}
        <a href="https://www.minouq.com" target="_blank" rel="noopener noreferrer" style={{ color: COLORS.seal, fontWeight: 600 }}>
          www.minouq.com
        </a>
      </div>
    </div>
  );
}

function LegalPage({ title, updated, sections }) {
  return (
    <div style={{ padding: "20px 4px", maxWidth: 640, margin: "0 auto" }}>
      <div style={{ fontSize: 22, fontWeight: 700, color: COLORS.ink, marginBottom: 4, textAlign: "center" }}>{title}</div>
      <div style={{ fontSize: 12, color: COLORS.metadata, textAlign: "center", marginBottom: 24 }}>Last updated: {updated}</div>
      {sections.map((s, i) => (
        <div key={i} style={{ marginBottom: 20 }}>
          <div style={{ fontSize: 14.5, fontWeight: 600, color: COLORS.ink, marginBottom: 6 }}>{s.heading}</div>
          <div style={{ fontSize: 14, color: COLORS.inkSoft, lineHeight: 1.7, whiteSpace: "pre-line", textAlign: "justify" }}>{s.body}</div>
        </div>
      ))}
    </div>
  );
}

function PrivacyTab() {
  return (
    <LegalPage
      title="Privacy Policy"
      updated="September 2026"
      sections={[
        {
          heading: "1. What we collect",
          body: "When you create an account, we collect your email address and password (handled securely by our authentication provider, Supabase — we never see or store your password in plain text). If you use the character/word lookup feature, we send the character or word you enter to Anthropic's Claude API to generate meaning, pinyin, and Sino-Vietnamese data. If you submit feedback, we store the message you write and, optionally, an email address you choose to provide.",
        },
        {
          heading: "2. How we use it",
          body: "We use your account data to let you save your own custom characters, words, and lists, track your flashcard progress, and manage your subscription tier. We use submitted feedback to improve the app. We do not sell your personal data to third parties.",
        },
        {
          heading: "3. Third-party services",
          body: "We use Supabase for authentication and data storage, and Anthropic's API for character/word lookups. These providers process data on our behalf under their own privacy and security terms.",
        },
        {
          heading: "4. Cookies and local storage",
          body: "We use your browser's local storage to remember small preferences, such as whether you've already chosen a display language, so we don't ask again on every visit. We don't use tracking or advertising cookies.",
        },
        {
          heading: "5. Your choices",
          body: "You can edit or delete your own custom characters, words, and lists at any time from within the app. To delete your account entirely or request a copy of your data, contact us using the details below.",
        },
        {
          heading: "6. Children's privacy",
          body: "This app is intended for general audiences learning Chinese. We do not knowingly collect personal data from children without appropriate parental consent where required by local law.",
        },
        {
          heading: "7. Changes to this policy",
          body: "We may update this policy as the app evolves. We'll update the date above when we do.",
        },
        {
          heading: "8. Contact",
          body: "Questions about this policy? Reach us at hello@minouq.com.",
        },
      ]}
    />
  );
}

function TermsTab() {
  return (
    <LegalPage
      title="Terms of Service"
      updated="September 2026"
      sections={[
        {
          heading: "1. Using this service",
          body: "MinouQ is a tool for learning Chinese characters through their component structure. By creating an account or using the app, you agree to these terms.",
        },
        {
          heading: "2. Accounts",
          body: "You're responsible for keeping your login credentials secure and for activity that happens under your account. You must provide a valid email address to create an account.",
        },
        {
          heading: "3. Subscription tiers and lookups",
          body: "Free accounts include a limited number of automatic character/word lookups per period. Paid tiers unlock higher or unlimited lookup allowances and access to additional content. Pricing and tier details are shown on the Pricing page.",
        },
        {
          heading: "4. Acceptable use",
          body: "Please don't use the lookup feature to abuse, overload, or attempt to extract data from our systems beyond normal use, and don't submit content that's illegal, abusive, or infringes on others' rights.",
        },
        {
          heading: "5. Your content",
          body: "Characters, words, and lists you add remain associated with your account. By submitting feedback, you allow us to use it to improve the app.",
        },
        {
          heading: "6. Changes and availability",
          body: "We may add, change, or remove features over time, and may suspend accounts that violate these terms. We aim for reliable uptime but don't guarantee uninterrupted service.",
        },
        {
          heading: "7. Disclaimer",
          body: "This app is provided \"as is.\" While we aim for accuracy, character meanings, pinyin, and readings — including those generated automatically — may occasionally contain errors; please cross-check important learning material against a trusted dictionary.",
        },
        {
          heading: "8. Contact",
          body: "Questions about these terms? Reach us at hello@minouq.com.",
        },
      ]}
    />
  );
}

function FeedbackTab({ meaningDisplay, userId }) {
  const [text, setText] = useState("");
  const [email, setEmail] = useState("");
  const [status, setStatus] = useState("idle"); // idle | sending | done | error

  async function handleSubmit(e) {
    e.preventDefault();
    if (!text.trim()) {
      setStatus("empty");
      return;
    }
    setStatus("sending");
    try {
      const { error } = await supabase.from("feedback").insert({
        message: text.trim(),
        email: email.trim() || null,
        user_id: userId || null,
      });
      if (error) throw error;
      setStatus("done");
      setText("");
      setEmail("");
    } catch (err) {
      console.error("Feedback submit failed:", err);
      setStatus("error");
    }
  }

  return (
    <div style={{ maxWidth: 480, margin: "0 auto", padding: "20px 4px", textAlign: "center" }}>
      <div style={{ fontSize: 22, fontWeight: 700, color: COLORS.ink, marginBottom: 8 }}>
        {t("feedback_title", meaningDisplay)}
      </div>
      <div style={{ fontSize: 14, color: COLORS.inkSoft, marginBottom: 6, lineHeight: 1.6 }}>
        {t("feedback_body", meaningDisplay)}
      </div>
      <div style={{ fontSize: 13, color: COLORS.metadata, marginBottom: 20, lineHeight: 1.5 }}>
        {t("feedback_contact_line", meaningDisplay)}
      </div>
      <form onSubmit={handleSubmit}>
        <textarea
          value={text}
          onChange={(e) => {
            setText(e.target.value);
            if (status === "empty" || status === "error") setStatus("idle");
          }}
          placeholder={t("feedback_placeholder", meaningDisplay)}
          rows={5}
          style={{ ...inputStyle, width: "100%", boxSizing: "border-box", resize: "vertical", marginBottom: 10, fontFamily: "'Noto Sans', sans-serif" }}
        />
        <input
          type="email"
          value={email}
          onChange={(e) => setEmail(e.target.value)}
          placeholder={t("feedback_email_placeholder", meaningDisplay)}
          style={{ ...inputStyle, width: "100%", boxSizing: "border-box", marginBottom: 12 }}
        />
        <button type="submit" className="seal-btn" style={{ ...sealBtnStyle, padding: "10px 26px", fontSize: 14 }} disabled={status === "sending"}>
          {status === "sending" ? t("feedback_sending", meaningDisplay) : t("feedback_submit", meaningDisplay)}
        </button>
      </form>
      {status === "done" && (
        <div style={{ marginTop: 14, fontSize: 13.5, fontWeight: 600, color: COLORS.seal }}>{t("feedback_success", meaningDisplay)}</div>
      )}
      {status === "error" && (
        <div style={{ marginTop: 14, fontSize: 13.5, fontWeight: 600, color: COLORS.error }}>{t("feedback_error", meaningDisplay)}</div>
      )}
      {status === "empty" && (
        <div style={{ marginTop: 14, fontSize: 13.5, fontWeight: 600, color: COLORS.error }}>{t("feedback_empty", meaningDisplay)}</div>
      )}
    </div>
  );
}

function PremiumTab({ meaningDisplay }) {
  const tiers = ["Free", "Premium"];
  const lookups = { Free: "1,000", Premium: "5,000" };
  const ads = {
    Free: t("pricing_yes", meaningDisplay),
    Premium: t("pricing_no", meaningDisplay),
  };
  const rows = [
    { label: t("pricing_row_lookups", meaningDisplay), values: lookups },
    { label: t("pricing_row_ads", meaningDisplay), values: ads },
    { label: t("pricing_row_rate", meaningDisplay), values: Object.fromEntries(tiers.map((tier) => [tier, tier === "Free" ? t("pricing_free_rate", meaningDisplay) : t("pricing_tbd", meaningDisplay)])) },
  ];

  const tierColors = { Free: COLORS.metadata, Premium: COLORS.seal };
  const cellStyle = { padding: "10px 12px", fontSize: 12.5, color: COLORS.inkSoft, textAlign: "center", borderBottom: `1px solid ${COLORS.hairline}` };
  const headStyle = { ...cellStyle, fontWeight: 700, fontSize: 13, borderBottom: "none" };
  const labelCellStyle = { ...cellStyle, textAlign: "left", fontWeight: 600, color: COLORS.ink };

  return (
    <div style={{ maxWidth: 620, margin: "0 auto" }}>
      <div style={{ fontSize: 22, fontWeight: 700, color: COLORS.ink, marginBottom: 14, textAlign: "center" }}>
        {t("tab_premium", meaningDisplay)}
      </div>
      <div style={{ fontSize: 14, color: COLORS.inkSoft, textAlign: "justify", marginBottom: 26, lineHeight: 1.7 }}>
        {t("pricing_message", meaningDisplay)}
      </div>

      <div style={{ overflowX: "auto", marginBottom: 20 }}>
        <table style={{ width: "100%", borderCollapse: "collapse", maxWidth: 380, margin: "0 auto" }}>
          <thead>
            <tr>
              <th style={{ ...headStyle, textAlign: "left" }}></th>
              {tiers.map((tier) => (
                <th key={tier} style={{ ...headStyle, color: tierColors[tier], borderTop: `2px solid ${tierColors[tier]}`, paddingTop: 10 }}>
                  {tier}
                </th>
              ))}
            </tr>
          </thead>
          <tbody>
            {rows.map((row, i) => (
              <tr key={row.label} style={{ background: i % 2 === 1 ? COLORS.chipBg : "transparent" }}>
                <td style={labelCellStyle}>{row.label}</td>
                {tiers.map((tier) => {
                  const isAdsRow = row.label === t("pricing_row_ads", meaningDisplay);
                  const isYes = row.values[tier] === t("pricing_yes", meaningDisplay);
                  return (
                    <td
                      key={tier}
                      style={{
                        ...cellStyle,
                        fontWeight: isAdsRow ? 600 : 400,
                        color: isAdsRow ? (isYes ? "#C62828" : "#2E7D32") : COLORS.inkSoft,
                      }}
                    >
                      {row.values[tier]}
                    </td>
                  );
                })}
              </tr>
            ))}
          </tbody>
        </table>
      </div>

      <div style={{ background: COLORS.card, border: `1px solid ${COLORS.hairline}`, borderRadius: 12, padding: "16px 18px", marginBottom: 24 }}>
        <div style={{ fontSize: 12, fontWeight: 700, color: COLORS.ink, textTransform: "uppercase", letterSpacing: 0.5, marginBottom: 8 }}>
          {t("pricing_lookup_explainer_title", meaningDisplay)}
        </div>
        <div style={{ fontSize: 12.5, color: COLORS.inkSoft, lineHeight: 1.6, textAlign: "justify" }}>
          {t("pricing_lookup_explainer_body", meaningDisplay)}
        </div>
      </div>

      <div style={{ fontSize: 12, fontWeight: 700, color: COLORS.ink, textTransform: "uppercase", letterSpacing: 0.5, marginBottom: 10 }}>
        {t("pricing_courses_title", meaningDisplay)}
      </div>
      <div style={{ fontSize: 12.5, color: COLORS.inkSoft, lineHeight: 1.6, textAlign: "justify" }}>
        {t("pricing_courses_body", meaningDisplay)}
      </div>
    </div>
  );
}

function AdminPanel({ isAdmin, allListNamesInUse, meaningDisplay, characterList, wordList, bushouList, decks, onDecksChanged, userId, adminBadges }) {
  const [users, setUsers] = useState([]);
  const [loading, setLoading] = useState(true);
  const [query, setQuery] = useState("");
  const [tierFilter, setTierFilter] = useState("Tất cả");
  const [statusFilter, setStatusFilter] = useState("all"); // all | disabled | enabled
  const [courseFilter, setCourseFilter] = useState("Tất cả");
  const [editingId, setEditingId] = useState(null);
  const [editTier, setEditTier] = useState("Free");
  const [editLimit, setEditLimit] = useState("100");
  const [editCourseName, setEditCourseName] = useState("");
  const [message, setMessage] = useState(null);

  // List access management
  const [listSettings, setListSettings] = useState([]);
  const [listCourseAccess, setListCourseAccess] = useState([]);
  const [listsLoading, setListsLoading] = useState(true);
  const [editingListName, setEditingListName] = useState(null);
  const [editAdminOnly, setEditAdminOnly] = useState(false);
  const [editAllowedTiers, setEditAllowedTiers] = useState([]);
  const [newCourseGrant, setNewCourseGrant] = useState("");
  const [listMessage, setListMessage] = useState(null);

  // Feedback review
  const [feedbackList, setFeedbackList] = useState([]);
  const [feedbackLoading, setFeedbackLoading] = useState(true);

  // Blog post management
  const [blogPosts, setBlogPosts] = useState([]);
  const [blogLoading, setBlogLoading] = useState(true);
  const [editingBlogId, setEditingBlogId] = useState(null); // null = not editing, "new" = creating, or a post id
  const [blogTitle, setBlogTitle] = useState("");
  const [blogBody, setBlogBody] = useState("");
  const [blogCategory, setBlogCategory] = useState("news");
  const [blogLink, setBlogLink] = useState("");
  const [blogPublished, setBlogPublished] = useState(false);
  const [blogMessage, setBlogMessage] = useState(null);

  useEffect(() => {
    if (!isAdmin) return;
    loadUsers();
    loadListSettings();
    loadFeedback();
    loadBlogPosts();
    loadSuggestions();
    loadComments();
    loadMsgThreads();
    // eslint-disable-next-line react-hooks/exhaustive-deps
  }, [isAdmin]);

  async function loadBlogPosts() {
    setBlogLoading(true);
    const { data, error } = await supabase
      .from("blog_posts")
      .select("id, title, body, category, external_link, published, created_at")
      .order("created_at", { ascending: false });
    if (!error) setBlogPosts(data || []);
    setBlogLoading(false);
  }

  function startNewBlogPost() {
    setEditingBlogId("new");
    setBlogTitle("");
    setBlogBody("");
    setBlogCategory("news");
    setBlogLink("");
    setBlogPublished(false);
    setBlogMessage(null);
  }

  function startEditBlogPost(post) {
    setEditingBlogId(post.id);
    setBlogTitle(post.title);
    setBlogBody(post.body);
    setBlogCategory(post.category);
    setBlogLink(post.external_link || "");
    setBlogPublished(post.published);
    setBlogMessage(null);
  }

  async function saveBlogPost() {
    if (!blogTitle.trim() || !blogBody.trim()) {
      setBlogMessage({ type: "error", text: t("admin_blog_need_title_body", meaningDisplay) });
      return;
    }
    const payload = {
      title: blogTitle.trim(),
      body: blogBody.trim(),
      category: blogCategory,
      external_link: blogLink.trim() || null,
      published: blogPublished,
      updated_at: new Date().toISOString(),
    };
    let error;
    if (editingBlogId === "new") {
      ({ error } = await supabase.from("blog_posts").insert(payload));
    } else {
      ({ error } = await supabase.from("blog_posts").update(payload).eq("id", editingBlogId));
    }
    if (error) {
      setBlogMessage({ type: "error", text: error.message });
      return;
    }
    setEditingBlogId(null);
    await loadBlogPosts();
  }

  async function deleteBlogPost(id) {
    if (!window.confirm(t("admin_blog_confirm_delete", meaningDisplay))) return;
    const { error } = await supabase.from("blog_posts").delete().eq("id", id);
    if (!error) setBlogPosts((prev) => prev.filter((p) => p.id !== id));
  }

  async function toggleBlogPublish(post) {
    const nextPublished = !post.published;
    const { error } = await supabase
      .from("blog_posts")
      .update({ published: nextPublished, updated_at: new Date().toISOString() })
      .eq("id", post.id);
    if (!error) {
      setBlogPosts((prev) => prev.map((p) => (p.id === post.id ? { ...p, published: nextPublished } : p)));
    }
  }

  // Blog comment moderation
  const [comments, setComments] = useState([]);
  const [commentsLoading, setCommentsLoading] = useState(true);
  const [commentStatusFilter, setCommentStatusFilter] = useState("pending");

  async function loadComments() {
    setCommentsLoading(true);
    const [commentsRes, postsRes] = await Promise.all([
      supabase.from("blog_comments").select("*").order("created_at", { ascending: false }),
      supabase.from("blog_posts").select("id, title"),
    ]);
    const postTitles = new Map((postsRes.data || []).map((p) => [p.id, p.title]));
    setComments((commentsRes.data || []).map((c) => ({ ...c, postTitle: postTitles.get(c.post_id) || null })));
    setCommentsLoading(false);
  }

  async function setCommentStatus(id, status) {
    const { error } = await supabase.from("blog_comments").update({ status }).eq("id", id);
    if (!error) setComments((prev) => prev.map((c) => (c.id === id ? { ...c, status } : c)));
  }

  async function deleteComment(id) {
    if (!window.confirm(t("admin_comment_confirm_delete", meaningDisplay))) return;
    const { error } = await supabase.from("blog_comments").delete().eq("id", id);
    if (!error) setComments((prev) => prev.filter((c) => c.id !== id));
  }

  // Admin messaging inbox
  const [msgThreads, setMsgThreads] = useState(null); // null = loading
  const [msgSelectedUser, setMsgSelectedUser] = useState(null); // { user_id, email }
  const [msgThread, setMsgThread] = useState(null);
  const [msgDraft, setMsgDraft] = useState("");
  const [msgSending, setMsgSending] = useState(false);
  const [msgSearchQuery, setMsgSearchQuery] = useState("");
  const [msgUsersList, setMsgUsersList] = useState([]);

  useEffect(() => {
    (async () => {
      const { data } = await supabase.from("profiles").select("user_id, email").order("email");
      setMsgUsersList(data || []);
    })();
  }, []);

  async function loadMsgThreads() {
    setMsgThreads(null);
    const [messagesRes, usersRes] = await Promise.all([
      supabase.from("messages").select("*").order("created_at", { ascending: false }),
      supabase.from("profiles").select("user_id, email"),
    ]);
    const emailByUser = new Map((usersRes.data || []).map((u) => [u.user_id, u.email]));
    const byUser = new Map();
    (messagesRes.data || []).forEach((m) => {
      if (!byUser.has(m.user_id)) {
        byUser.set(m.user_id, { user_id: m.user_id, email: emailByUser.get(m.user_id) || m.user_id, latest: m, unread: 0 });
      }
      if (m.sender === "user" && !m.read) byUser.get(m.user_id).unread += 1;
    });
    setMsgThreads(Array.from(byUser.values()).sort((a, b) => new Date(b.latest.created_at) - new Date(a.latest.created_at)));
  }

  const msgSearchResults =
    msgSearchQuery.trim().length > 0
      ? msgUsersList.filter((u) => (u.email || "").toLowerCase().includes(msgSearchQuery.trim().toLowerCase())).slice(0, 8)
      : [];

  async function selectMsgThread(user) {
    setMsgSelectedUser(user);
    setMsgSearchQuery("");
    setMsgThread(null);
    const { data, error } = await supabase.from("messages").select("*").eq("user_id", user.user_id).order("created_at", { ascending: true });
    if (!error) {
      setMsgThread(data || []);
      const unreadIds = (data || []).filter((m) => m.sender === "user" && !m.read).map((m) => m.id);
      if (unreadIds.length > 0) {
        await supabase.from("messages").update({ read: true }).in("id", unreadIds);
        loadMsgThreads();
      }
    }
  }

  async function sendMsgReply(e) {
    e.preventDefault();
    if (!msgDraft.trim() || msgSending || !msgSelectedUser) return;
    setMsgSending(true);
    const { error } = await supabase.from("messages").insert({ user_id: msgSelectedUser.user_id, sender: "admin", message: msgDraft.trim() });
    if (!error) {
      setMsgDraft("");
      const { data } = await supabase.from("messages").select("*").eq("user_id", msgSelectedUser.user_id).order("created_at", { ascending: true });
      setMsgThread(data || []);
      loadMsgThreads();
    }
    setMsgSending(false);
  }

  // Deck management
  const [editingDeckId, setEditingDeckId] = useState(null); // null = not editing, "new" = creating, or a deck id
  const [deckName, setDeckName] = useState("");
  const [deckDescription, setDeckDescription] = useState("");
  const [deckLists, setDeckLists] = useState([]); // [{content_type, list_name}]
  const [deckAddType, setDeckAddType] = useState("char");
  const [deckAddList, setDeckAddList] = useState("");
  const [deckMessage, setDeckMessage] = useState(null);

  const availableListsByType = {
    char: Array.from(new Set((characterList || []).flatMap((c) => getLists(c)))).sort((a, b) => a.localeCompare(b, "vi")),
    word: Array.from(new Set((wordList || []).flatMap((w) => w.lists || []))).sort((a, b) => a.localeCompare(b, "vi")),
    bushou: Array.from(new Set((bushouList || []).flatMap((b) => b.lists || []))).sort((a, b) => {
      const numA = parseInt(a, 10);
      const numB = parseInt(b, 10);
      if (!isNaN(numA) && !isNaN(numB)) return numA - numB;
      return a.localeCompare(b, "vi");
    }),
  };

  function startNewDeck() {
    setEditingDeckId("new");
    setDeckName("");
    setDeckDescription("");
    setDeckLists([]);
    setDeckAddType("char");
    setDeckAddList("");
    setDeckMessage(null);
  }

  function startEditDeck(deck) {
    setEditingDeckId(deck.id);
    setDeckName(deck.name);
    setDeckDescription(deck.description || "");
    setDeckLists(deck.lists || []);
    setDeckAddType("char");
    setDeckAddList("");
    setDeckMessage(null);
  }

  function addDeckList() {
    if (!deckAddList) return;
    if (deckLists.some((l) => l.content_type === deckAddType && l.list_name === deckAddList)) return;
    setDeckLists([...deckLists, { content_type: deckAddType, list_name: deckAddList }]);
    setDeckAddList("");
  }

  function removeDeckList(contentType, listName) {
    setDeckLists(deckLists.filter((l) => !(l.content_type === contentType && l.list_name === listName)));
  }

  async function saveDeck() {
    if (!deckName.trim()) {
      setDeckMessage({ type: "error", text: t("admin_deck_need_name", meaningDisplay) });
      return;
    }
    let deckId = editingDeckId;
    if (editingDeckId === "new") {
      const { data, error } = await supabase
        .from("decks")
        .insert({ name: deckName.trim(), description: deckDescription.trim() || null })
        .select()
        .single();
      if (error) {
        setDeckMessage({ type: "error", text: error.message });
        return;
      }
      deckId = data.id;
    } else {
      const { error } = await supabase
        .from("decks")
        .update({ name: deckName.trim(), description: deckDescription.trim() || null })
        .eq("id", editingDeckId);
      if (error) {
        setDeckMessage({ type: "error", text: error.message });
        return;
      }
      await supabase.from("deck_lists").delete().eq("deck_id", editingDeckId);
    }
    if (deckLists.length > 0) {
      const { error } = await supabase
        .from("deck_lists")
        .insert(deckLists.map((l) => ({ deck_id: deckId, content_type: l.content_type, list_name: l.list_name })));
      if (error) {
        setDeckMessage({ type: "error", text: error.message });
        return;
      }
    }
    setEditingDeckId(null);
    if (onDecksChanged) onDecksChanged();
  }

  async function deleteDeck(id) {
    if (!window.confirm(t("admin_deck_confirm_delete", meaningDisplay))) return;
    const { error } = await supabase.from("decks").delete().eq("id", id);
    if (!error && onDecksChanged) onDecksChanged();
  }

  // User library browsing & copying
  const [libSearchQuery, setLibSearchQuery] = useState("");
  const [libSelectedUser, setLibSelectedUser] = useState(null); // { user_id, email }
  const [libUserChars, setLibUserChars] = useState(null);
  const [libUserWords, setLibUserWords] = useState(null);
  const [libUserBushou, setLibUserBushou] = useState(null);
  const [libUserDecks, setLibUserDecks] = useState(null);
  const [libLoading, setLibLoading] = useState(false);
  const [libCopyMessage, setLibCopyMessage] = useState(null);
  const [libOnlyNew, setLibOnlyNew] = useState(false);
  const [libUsersList, setLibUsersList] = useState([]);

  // What the admin already has -- used to flag genuinely new items in a
  // browsed user's library, so new contributions are easy to spot at a
  // glance rather than needing to re-check each one manually.
  const myCharKeys = useMemo(() => new Set((characterList || []).map((c) => c.char)), [characterList]);
  const myWordKeys = useMemo(() => new Set((wordList || []).map((w) => w.word)), [wordList]);
  const myBushouKeys = useMemo(() => new Set((bushouList || []).map((b) => b.char)), [bushouList]);
  function isNewToMe(type, item) {
    const key = item.char || item.word;
    if (type === "char") return !myCharKeys.has(key);
    if (type === "word") return !myWordKeys.has(key);
    return !myBushouKeys.has(key);
  }

  useEffect(() => {
    // A small dedicated user search for this feature, separate from the
    // Users tab's own state, so the two sections don't interfere.
    (async () => {
      const { data } = await supabase.from("profiles").select("user_id, email").order("email");
      setLibUsersList(data || []);
    })();
  }, []);

  const libSearchResults =
    libSearchQuery.trim().length > 0
      ? libUsersList.filter((u) => (u.email || "").toLowerCase().includes(libSearchQuery.trim().toLowerCase())).slice(0, 8)
      : [];

  async function selectLibUser(user) {
    setLibSelectedUser(user);
    setLibSearchQuery("");
    setLibCopyMessage(null);
    setLibLoading(true);
    const [charsRes, wordsRes, bushouRes, decksRes, deckListsRes] = await Promise.all([
      supabase.from("custom_characters").select("*").eq("user_id", user.user_id),
      supabase.from("custom_words").select("*").eq("user_id", user.user_id),
      supabase.from("custom_bushou").select("*").eq("user_id", user.user_id),
      supabase.from("decks").select("*").eq("user_id", user.user_id),
      supabase.from("deck_lists").select("*"),
    ]);
    setLibUserChars((charsRes.data || []).map(rowToChar));
    setLibUserWords((wordsRes.data || []).map(rowToWord));
    setLibUserBushou((bushouRes.data || []).map(rowToBushou));
    const byDeck = new Map();
    (deckListsRes.data || []).forEach((row) => {
      if (!byDeck.has(row.deck_id)) byDeck.set(row.deck_id, []);
      byDeck.get(row.deck_id).push({ content_type: row.content_type, list_name: row.list_name });
    });
    setLibUserDecks((decksRes.data || []).map((d) => ({ id: d.id, name: d.name, lists: byDeck.get(d.id) || [] })));
    setLibLoading(false);
  }

  function exitLibView() {
    setLibSelectedUser(null);
    setLibUserChars(null);
    setLibUserWords(null);
    setLibUserBushou(null);
    setLibUserDecks(null);
    setLibCopyMessage(null);
  }

  async function copyOneItem(contentType, item) {
    const table = contentType === "char" ? "custom_characters" : contentType === "word" ? "custom_words" : "custom_bushou";
    const toRow = contentType === "char" ? charToRow : contentType === "word" ? wordToRow : bushouToRow;
    const row = {
      ...toRow(item, userId),
      copied_from_user_id: libSelectedUser.user_id,
      copied_from_email: libSelectedUser.email,
    };
    const { error } = await supabase.from(table).upsert(row, { onConflict: contentType === "word" ? "user_id,word" : "user_id,char" });
    return !error;
  }

  async function copyItem(contentType, item) {
    const ok = await copyOneItem(contentType, item);
    setLibCopyMessage({ type: ok ? "success" : "error", text: ok ? t("admin_lib_copy_done", meaningDisplay) : t("admin_lib_copy_error", meaningDisplay) });
  }

  async function copyList(contentType, listName) {
    const source = contentType === "char" ? libUserChars : contentType === "word" ? libUserWords : libUserBushou;
    const matches = (source || []).filter((item) =>
      contentType === "word" ? (item.lists || []).includes(listName) : (item.lists || []).includes(listName)
    );
    let successCount = 0;
    for (const item of matches) {
      if (await copyOneItem(contentType, item)) successCount += 1;
    }
    setLibCopyMessage({ type: "success", text: t("admin_lib_copy_list_done", meaningDisplay, successCount, matches.length) });
  }

  async function copyDeck(deck) {
    const seen = new Set();
    let copiedCount = 0;
    for (const { content_type, list_name } of deck.lists || []) {
      const source = content_type === "char" ? libUserChars : content_type === "word" ? libUserWords : libUserBushou;
      const matches = (source || []).filter((item) => (item.lists || []).includes(list_name));
      for (const item of matches) {
        const key = `${content_type}:${item.char || item.word}`;
        if (seen.has(key)) continue;
        seen.add(key);
        if (await copyOneItem(content_type, item)) copiedCount += 1;
      }
    }
    const { error } = await supabase.from("decks").insert({
      name: deck.name,
      user_id: userId,
      copied_from_user_id: libSelectedUser.user_id,
      copied_from_email: libSelectedUser.email,
    });
    if (!error) {
      const { data: newDeck } = await supabase.from("decks").select("id").eq("user_id", userId).eq("name", deck.name).order("created_at", { ascending: false }).limit(1).single();
      if (newDeck && deck.lists.length > 0) {
        await supabase.from("deck_lists").insert(deck.lists.map((l) => ({ deck_id: newDeck.id, content_type: l.content_type, list_name: l.list_name })));
      }
    }
    if (onDecksChanged) onDecksChanged();
    setLibCopyMessage({ type: error ? "error" : "success", text: error ? t("admin_lib_copy_error", meaningDisplay) : t("admin_lib_copy_deck_done", meaningDisplay, copiedCount) });
  }

  // Card revision suggestions
  const [suggestions, setSuggestions] = useState([]);
  const [suggestionsLoading, setSuggestionsLoading] = useState(true);
  const [suggestionStatusFilter, setSuggestionStatusFilter] = useState("new");

  async function loadSuggestions() {
    setSuggestionsLoading(true);
    const { data, error } = await supabase.from("card_suggestions").select("*").order("created_at", { ascending: false });
    if (!error) setSuggestions(data || []);
    setSuggestionsLoading(false);
  }

  function findSuggestionItem(s) {
    const list = s.content_type === "char" ? characterList : s.content_type === "word" ? wordList : bushouList;
    if (s.content_type === "word") return (list || []).find((w) => w.word === s.item_key);
    return (list || []).find((x) => x.char === s.item_key);
  }

  async function setSuggestionStatus(id, status) {
    const { error } = await supabase.from("card_suggestions").update({ status }).eq("id", id);
    if (!error) setSuggestions((prev) => prev.map((s) => (s.id === id ? { ...s, status } : s)));
  }

  async function deleteSuggestion(id) {
    if (!window.confirm(t("admin_suggestion_confirm_delete", meaningDisplay))) return;
    const { error } = await supabase.from("card_suggestions").delete().eq("id", id);
    if (!error) setSuggestions((prev) => prev.filter((s) => s.id !== id));
  }

  async function loadFeedback() {
    setFeedbackLoading(true);
    const { data, error } = await supabase
      .from("feedback")
      .select("id, message, email, user_id, read, created_at")
      .order("created_at", { ascending: false });
    if (error) {
      setFeedbackLoading(false);
      return;
    }
    const rows = data || [];
    const userIds = [...new Set(rows.filter((f) => f.user_id).map((f) => f.user_id))];
    let accountEmails = {};
    if (userIds.length > 0) {
      const { data: profileRows } = await supabase.from("profiles").select("user_id, email").in("user_id", userIds);
      (profileRows || []).forEach((p) => {
        accountEmails[p.user_id] = p.email;
      });
    }
    setFeedbackList(rows.map((f) => ({ ...f, accountEmail: f.user_id ? accountEmails[f.user_id] : null })));
    setFeedbackLoading(false);
  }

  async function markFeedbackRead(id) {
    const { error } = await supabase.from("feedback").update({ read: true }).eq("id", id);
    if (!error) setFeedbackList((prev) => prev.map((f) => (f.id === id ? { ...f, read: true } : f)));
  }

  async function handleDeleteFeedback(id) {
    const { error } = await supabase.from("feedback").delete().eq("id", id);
    if (!error) setFeedbackList((prev) => prev.filter((f) => f.id !== id));
  }

  async function loadUsers() {
    setLoading(true);
    const { data, error } = await supabase
      .from("profiles")
      .select("user_id, email, is_admin, tier, lookup_count, lookup_limit, disabled, course_name")
      .order("email", { ascending: true });
    if (!error) setUsers(data || []);
    setLoading(false);
  }

  async function loadListSettings() {
    setListsLoading(true);
    const [lsRes, lcaRes] = await Promise.all([
      supabase.from("list_settings").select("*"),
      supabase.from("list_course_access").select("*"),
    ]);
    if (!lsRes.error) setListSettings(lsRes.data || []);
    if (!lcaRes.error) setListCourseAccess(lcaRes.data || []);
    setListsLoading(false);
  }

  function getListSetting(name) {
    return listSettings.find((s) => s.name === name);
  }

  function getListCourseGrants(name) {
    return listCourseAccess.filter((g) => g.list_name === name).map((g) => g.course_name);
  }

  function startEditList(name) {
    const setting = getListSetting(name);
    setEditingListName(name);
    setEditAdminOnly(setting ? setting.admin_only : false);
    setEditAllowedTiers(setting && setting.allowed_tiers ? setting.allowed_tiers : []);
    setNewCourseGrant("");
  }

  function toggleEditTier(t) {
    setEditAllowedTiers((prev) => (prev.includes(t) ? prev.filter((x) => x !== t) : [...prev, t]));
  }

  async function saveListSetting(name) {
    const { error } = await supabase
      .from("list_settings")
      .upsert({ name, admin_only: editAdminOnly, allowed_tiers: editAllowedTiers }, { onConflict: "name" });
    if (error) {
      setListMessage({ type: "error", text: t("admin_save_failed", meaningDisplay, error.message) });
      return;
    }
    setListSettings((prev) => {
      const without = prev.filter((s) => s.name !== name);
      return [...without, { name, admin_only: editAdminOnly, allowed_tiers: editAllowedTiers }];
    });
    setEditingListName(null);
    setListMessage({ type: "success", text: t("admin_saved", meaningDisplay) });
    setTimeout(() => setListMessage(null), 2500);
  }

  async function addCourseGrant(name) {
    const course = newCourseGrant.trim();
    if (!course) return;
    // A course grant implies the list needs a list_settings row to exist
    // at all (otherwise it's already open to everyone and the grant is
    // moot) -- make sure one exists first.
    if (!getListSetting(name)) {
      const { error: upsertErr } = await supabase
        .from("list_settings")
        .upsert({ name, admin_only: false, allowed_tiers: [] }, { onConflict: "name" });
      if (upsertErr) {
        setListMessage({ type: "error", text: "Không thể lưu: " + upsertErr.message });
        return;
      }
      setListSettings((prev) => [...prev, { name, admin_only: false, allowed_tiers: [] }]);
    }
    const { error } = await supabase.from("list_course_access").upsert({ list_name: name, course_name: course });
    if (error) {
      setListMessage({ type: "error", text: "Không thể thêm khóa học: " + error.message });
      return;
    }
    setListCourseAccess((prev) => [...prev, { list_name: name, course_name: course }]);
    setNewCourseGrant("");
  }

  async function removeCourseGrant(name, course) {
    const { error } = await supabase
      .from("list_course_access")
      .delete()
      .eq("list_name", name)
      .eq("course_name", course);
    if (error) {
      setListMessage({ type: "error", text: "Không thể xóa: " + error.message });
      return;
    }
    setListCourseAccess((prev) => prev.filter((g) => !(g.list_name === name && g.course_name === course)));
  }

  function startEdit(u) {
    setEditingId(u.user_id);
    setEditTier(u.tier && ALL_TIERS.includes(u.tier) ? u.tier : "Free");
    setEditLimit(String(u.lookup_limit != null ? u.lookup_limit : 1000));
    setEditCourseName(u.course_name || "");
  }

  function handleTierChange(newTier) {
    setEditTier(newTier);
    if (TIER_PRESETS[newTier] != null) setEditLimit(String(TIER_PRESETS[newTier]));
    // Leave the limit as-is when switching to Enrolled Course -- there's no
    // fixed preset for it, it's set manually per course/student.
  }

  async function saveEdit(userId) {
    const limitNum = parseInt(editLimit, 10);
    if (!Number.isFinite(limitNum) || limitNum < 0) {
      setMessage({ type: "error", text: t("admin_invalid_limit", meaningDisplay) });
      return;
    }
    const isCourse = editTier === "Enrolled Course";
    if (isCourse && !editCourseName.trim()) {
      setMessage({ type: "error", text: t("admin_need_course_name", meaningDisplay) });
      return;
    }
    const courseNameToSave = isCourse ? editCourseName.trim() : null;
    const { error } = await supabase
      .from("profiles")
      .update({ tier: editTier, lookup_limit: limitNum, course_name: courseNameToSave })
      .eq("user_id", userId);
    if (error) {
      setMessage({ type: "error", text: "Không thể lưu: " + error.message });
      return;
    }
    setUsers((prev) =>
      prev.map((u) => (u.user_id === userId ? { ...u, tier: editTier, lookup_limit: limitNum, course_name: courseNameToSave } : u))
    );
    setEditingId(null);
    setMessage({ type: "success", text: "Đã lưu." });
    setTimeout(() => setMessage(null), 2500);
  }

  async function resetUsage(userId) {
    if (!window.confirm(t("admin_confirm_reset_usage", meaningDisplay))) return;
    const { error } = await supabase.from("profiles").update({ lookup_count: 0 }).eq("user_id", userId);
    if (error) {
      setMessage({ type: "error", text: t("admin_reset_failed", meaningDisplay, error.message) });
      return;
    }
    setUsers((prev) => prev.map((u) => (u.user_id === userId ? { ...u, lookup_count: 0 } : u)));
    setMessage({ type: "success", text: t("admin_reset_done", meaningDisplay) });
    setTimeout(() => setMessage(null), 2500);
  }

  async function toggleDisabled(u) {
    const next = !u.disabled;
    if (next && !window.confirm(t("admin_confirm_disable", meaningDisplay, u.email || u.user_id))) return;
    const { error } = await supabase.from("profiles").update({ disabled: next }).eq("user_id", u.user_id);
    if (error) {
      setMessage({ type: "error", text: t("admin_update_failed", meaningDisplay, error.message) });
      return;
    }
    setUsers((prev) => prev.map((x) => (x.user_id === u.user_id ? { ...x, disabled: next } : x)));
    setMessage({ type: "success", text: next ? t("admin_disabled_done", meaningDisplay) : t("admin_enabled_done", meaningDisplay) });
    setTimeout(() => setMessage(null), 2500);
  }

  if (!isAdmin) return null;

  const existingCourseNames = Array.from(new Set(users.map((u) => u.course_name).filter(Boolean))).sort((a, b) =>
    a.localeCompare(b, "vi")
  );

  const filtered = users.filter((u) => {
    if (tierFilter !== "Tất cả" && (u.tier || "Free") !== tierFilter) return false;
    if (statusFilter === "disabled" && !u.disabled) return false;
    if (statusFilter === "enabled" && u.disabled) return false;
    if (courseFilter !== "Tất cả" && (u.course_name || "") !== courseFilter) return false;
    const q = query.trim().toLowerCase();
    if (!q) return true;
    return (u.email || "").toLowerCase().includes(q) || u.user_id.toLowerCase().includes(q);
  });

  const [adminSection, setAdminSection] = useState("users");
  const adminSections = [
    { id: "users", label: t("admin_nav_users", meaningDisplay) },
    { id: "lists", label: t("admin_nav_lists", meaningDisplay) },
    { id: "decks", label: t("admin_nav_decks", meaningDisplay) },
    { id: "userlib", label: t("admin_nav_userlib", meaningDisplay) },
    { id: "suggestions", label: t("admin_nav_suggestions", meaningDisplay), badge: adminBadges?.suggestions },
    { id: "feedback", label: t("admin_nav_feedback", meaningDisplay), badge: adminBadges?.feedback },
    { id: "blog", label: t("admin_nav_blog", meaningDisplay) },
    { id: "comments", label: t("admin_nav_comments", meaningDisplay), badge: adminBadges?.comments },
    { id: "messages", label: t("admin_nav_messages", meaningDisplay), badge: adminBadges?.messages },
  ];

  return (
    <div className="side-nav-layout" style={{ display: "flex", gap: 24, alignItems: "flex-start" }}>
      <div className="side-nav-menu" style={{ display: "flex", flexDirection: "column", gap: 4, width: 170, flexShrink: 0 }}>
        {adminSections.map((s) => (
          <button
            key={s.id}
            type="button"
            onClick={() => setAdminSection(s.id)}
            style={{
              textAlign: "left",
              background: adminSection === s.id ? COLORS.chipBg : "none",
              border: "none",
              borderLeft: `3px solid ${adminSection === s.id ? COLORS.seal : "transparent"}`,
              color: adminSection === s.id ? COLORS.ink : COLORS.inkSoft,
              fontWeight: adminSection === s.id ? 700 : 500,
              fontSize: 13,
              padding: "8px 12px",
              cursor: "pointer",
              borderRadius: 6,
              display: "flex",
              alignItems: "center",
              justifyContent: "space-between",
              gap: 8,
            }}
          >
            <span>{s.label}</span>
            {s.badge > 0 && (
              <span
                style={{
                  minWidth: 17,
                  height: 17,
                  borderRadius: 999,
                  background: COLORS.error,
                  color: "#FBF9EF",
                  fontSize: 10,
                  fontWeight: 700,
                  display: "flex",
                  alignItems: "center",
                  justifyContent: "center",
                  padding: "0 4px",
                  flexShrink: 0,
                }}
              >
                {s.badge > 9 ? "9+" : s.badge}
              </span>
            )}
          </button>
        ))}
      </div>
      <div style={{ flex: 1, minWidth: 0 }}>
      {adminSection === "users" && (
      <div>
      <div style={{ fontSize: 12.5, fontWeight: 600, color: COLORS.gold, marginBottom: 12, textTransform: "uppercase", letterSpacing: 0.8, textAlign: "center" }}>
        {t("admin_users_title", meaningDisplay)}
      </div>

      <div style={{ display: "flex", justifyContent: "center", flexWrap: "wrap", gap: 8, marginBottom: 10 }}>
        <input
          value={query}
          onChange={(e) => setQuery(e.target.value)}
          placeholder={t("admin_search_email_placeholder", meaningDisplay)}
          style={{ ...inputStyle, width: 220, textAlign: "center" }}
        />
        <select value={tierFilter} onChange={(e) => setTierFilter(e.target.value)} style={{ ...selectStyle, width: 150, flex: "none" }}>
          <option value="Tất cả" style={{ background: COLORS.chipBg, color: COLORS.ink, fontWeight: 700 }}>{t("admin_all_tiers", meaningDisplay)}</option>
          {ALL_TIERS.map((t) => (
            <option key={t} value={t} style={{ background: COLORS.chipBg, color: COLORS.ink, fontWeight: 700 }}>
              {t}
            </option>
          ))}
        </select>
        <select value={statusFilter} onChange={(e) => setStatusFilter(e.target.value)} style={{ ...selectStyle, width: 170, flex: "none" }}>
          <option value="all" style={{ background: COLORS.chipBg, color: COLORS.ink, fontWeight: 700 }}>{t("admin_all_status", meaningDisplay)}</option>
          <option value="enabled" style={{ background: COLORS.chipBg, color: COLORS.ink, fontWeight: 700 }}>{t("admin_status_enabled", meaningDisplay)}</option>
          <option value="disabled" style={{ background: COLORS.chipBg, color: COLORS.ink, fontWeight: 700 }}>{t("admin_status_disabled", meaningDisplay)}</option>
        </select>
        {existingCourseNames.length > 0 && (
          <select value={courseFilter} onChange={(e) => setCourseFilter(e.target.value)} style={{ ...selectStyle, width: 170, flex: "none" }}>
            <option value="Tất cả" style={{ background: COLORS.chipBg, color: COLORS.ink, fontWeight: 700 }}>{t("admin_all_courses", meaningDisplay)}</option>
            {existingCourseNames.map((c) => (
              <option key={c} value={c} style={{ background: COLORS.chipBg, color: COLORS.ink, fontWeight: 700 }}>
                {c}
              </option>
            ))}
          </select>
        )}
        <button type="button" onClick={loadUsers} className="ghost-btn" style={{ ...ghostBtnStyle, padding: "8px 14px", fontSize: 12.5 }}>
          {t("admin_refresh", meaningDisplay)}
        </button>
      </div>

      <div style={{ fontSize: 11.5, color: COLORS.inkSoft, textAlign: "center", marginBottom: 14 }}>
        {t("admin_user_count", meaningDisplay, filtered.length, users.length)}
      </div>

      {message && (
        <div
          style={{
            textAlign: "center",
            fontSize: 12.5,
            fontWeight: 600,
            marginBottom: 12,
            color: message.type === "error" ? COLORS.error : COLORS.bamboo,
          }}
        >
          {message.text}
        </div>
      )}

      {loading ? (
        <div style={{ textAlign: "center", color: COLORS.inkSoft, padding: 30 }}>{t("loading", meaningDisplay)}</div>
      ) : (
        <div style={{ display: "flex", flexDirection: "column", gap: 8 }}>
          {filtered.map((u) => (
            <div
              key={u.user_id}
              style={{
                background: COLORS.card,
                border: `1px solid ${COLORS.hairline}`,
                borderRadius: 11,
                padding: "10px 14px",
              }}
            >
              {editingId === u.user_id ? (
                <div style={{ display: "flex", flexWrap: "wrap", alignItems: "center", gap: 8 }}>
                  <div style={{ flex: "1 1 200px", fontSize: 13, color: COLORS.ink, fontWeight: 600 }}>
                    {u.email || u.user_id}
                    {u.is_admin && <span style={{ marginLeft: 6, fontSize: 10.5, color: COLORS.gold }}>(admin)</span>}
                  </div>
                  <select value={editTier} onChange={(e) => handleTierChange(e.target.value)} style={{ ...selectStyle, width: 150 }}>
                    {ALL_TIERS.map((t) => (
                      <option key={t} value={t} style={{ background: COLORS.chipBg, color: COLORS.ink, fontWeight: 700 }}>
                        {t}
                      </option>
                    ))}
                  </select>
                  <input
                    type="number"
                    value={editLimit}
                    onChange={(e) => setEditLimit(e.target.value)}
                    style={{ ...inputStyle, width: 100 }}
                  />
                  {editTier === "Enrolled Course" && (
                    <>
                      <input
                        value={editCourseName}
                        onChange={(e) => setEditCourseName(e.target.value)}
                        placeholder={t("admin_course_name_placeholder", meaningDisplay)}
                        list="admin-course-names"
                        style={{ ...inputStyle, width: 170 }}
                      />
                      <datalist id="admin-course-names">
                        {existingCourseNames.map((c) => (
                          <option key={c} value={c} />
                        ))}
                      </datalist>
                    </>
                  )}
                  <button type="button" onClick={() => saveEdit(u.user_id)} className="seal-btn" style={{ ...sealBtnStyle, padding: "6px 14px", fontSize: 12 }}>
                    {t("admin_save", meaningDisplay)}
                  </button>
                  <button
                    type="button"
                    onClick={() => setEditingId(null)}
                    className="ghost-btn"
                    style={{ ...ghostBtnStyle, padding: "6px 14px", fontSize: 12 }}
                  >
                    {t("admin_cancel", meaningDisplay)}
                  </button>
                </div>
              ) : (
                <div style={{ display: "grid", gridTemplateColumns: "minmax(160px, 1fr) 130px 90px auto", gap: 10, alignItems: "center" }}>
                  <div style={{ fontSize: 13, color: COLORS.ink, fontWeight: 600, overflow: "hidden", textOverflow: "ellipsis", whiteSpace: "nowrap" }} title={u.email || u.user_id}>
                    {u.email || u.user_id}
                    {u.is_admin && <span style={{ marginLeft: 6, fontSize: 10.5, color: COLORS.gold }}>(admin)</span>}
                    {u.disabled && (
                      <span
                        style={{
                          marginLeft: 6,
                          fontSize: 10,
                          fontWeight: 700,
                          padding: "2px 7px",
                          borderRadius: 999,
                          border: `1px solid ${COLORS.error}`,
                          background: "rgba(168,72,47,0.08)",
                          color: COLORS.error,
                        }}
                      >
                        {t("admin_disabled_badge", meaningDisplay)}
                      </span>
                    )}
                  </div>
                  <div style={{ fontSize: 12, color: COLORS.sealDark }}>
                    {u.tier || "Free"}
                    {u.tier === "Enrolled Course" && u.course_name && (
                      <span style={{ color: COLORS.inkSoft, fontWeight: 500 }}> · {u.course_name}</span>
                    )}
                  </div>
                  <div style={{ fontSize: 12, color: COLORS.inkSoft }}>
                    {u.lookup_count ?? 0} / {u.lookup_limit ?? 1000}
                  </div>
                  <div style={{ display: "flex", flexWrap: "wrap", gap: 6, justifyContent: "flex-end" }}>
                    <button
                      type="button"
                      onClick={() => startEdit(u)}
                      className="ghost-btn"
                      style={{ ...ghostBtnStyle, padding: "5px 10px", fontSize: 11.5 }}
                    >
                      {t("admin_edit", meaningDisplay)}
                    </button>
                    <button
                      type="button"
                      onClick={() => resetUsage(u.user_id)}
                      className="ghost-btn"
                      style={{ ...ghostBtnStyle, padding: "5px 10px", fontSize: 11.5 }}
                    >
                      {t("admin_reset_to_zero", meaningDisplay)}
                    </button>
                    {!u.is_admin && (
                      <button
                        type="button"
                        onClick={() => toggleDisabled(u)}
                        className="ghost-btn"
                        style={{
                          ...ghostBtnStyle,
                          padding: "5px 10px",
                          fontSize: 11.5,
                          borderColor: u.disabled ? COLORS.bamboo : COLORS.error,
                          color: u.disabled ? COLORS.bamboo : COLORS.error,
                        }}
                      >
                        {u.disabled ? t("admin_reenable", meaningDisplay) : t("admin_disable", meaningDisplay)}
                      </button>
                    )}
                  </div>
                </div>
              )}
            </div>
          ))}
          {filtered.length === 0 && (
            <div style={{ textAlign: "center", color: COLORS.inkSoft, padding: 20 }}>{t("admin_no_users_found", meaningDisplay)}</div>
          )}
        </div>
      )}
      </div>
      )}

      {adminSection === "lists" && (
      <div style={{ marginTop: 32, paddingTop: 22, borderTop: `1px dashed ${COLORS.grid}` }}>
        <div style={{ fontSize: 12.5, fontWeight: 600, color: COLORS.gold, marginBottom: 12, textTransform: "uppercase", letterSpacing: 0.8, textAlign: "center" }}>
          {t("admin_list_mgmt_title", meaningDisplay)}
        </div>
        <div style={{ fontSize: 11.5, color: COLORS.inkSoft, textAlign: "center", marginBottom: 16, lineHeight: 1.5 }}>
          {t("admin_list_mgmt_description", meaningDisplay)}
        </div>

        {listMessage && (
          <div
            style={{
              textAlign: "center",
              fontSize: 12.5,
              fontWeight: 600,
              marginBottom: 12,
              color: listMessage.type === "error" ? COLORS.error : COLORS.bamboo,
            }}
          >
            {listMessage.text}
          </div>
        )}

        {listsLoading ? (
          <div style={{ textAlign: "center", color: COLORS.inkSoft, padding: 20 }}>{t("loading", meaningDisplay)}</div>
        ) : (
          <div style={{ display: "flex", flexDirection: "column", gap: 8 }}>
            {allListNamesInUse.map((name) => {
              const setting = getListSetting(name);
              const courseGrants = getListCourseGrants(name);
              const isOpen = !setting;
              const isEditing = editingListName === name;
              return (
                <div
                  key={name}
                  style={{
                    background: COLORS.card,
                    border: `1px solid ${COLORS.hairline}`,
                    borderRadius: 11,
                    padding: "10px 14px",
                  }}
                >
                  {isEditing ? (
                    <div>
                      <div style={{ display: "flex", flexWrap: "wrap", alignItems: "center", gap: 10, marginBottom: 10 }}>
                        <div style={{ flex: "1 1 160px", fontSize: 13, color: COLORS.ink, fontWeight: 600 }}>{name}</div>
                        <label style={{ display: "flex", alignItems: "center", gap: 6, fontSize: 12.5, color: COLORS.ink, cursor: "pointer" }}>
                          <input type="checkbox" checked={editAdminOnly} onChange={(e) => setEditAdminOnly(e.target.checked)} />
                          {t("admin_admin_only_checkbox", meaningDisplay)}
                        </label>
                      </div>

                      {!editAdminOnly && (
                        <>
                          <div style={{ fontSize: 11, color: COLORS.inkSoft, marginBottom: 6 }}>{t("admin_allowed_tiers_label", meaningDisplay)}</div>
                          <div style={{ display: "flex", flexWrap: "wrap", gap: 6, marginBottom: 12 }}>
                            {ALL_TIERS.map((t) => (
                              <button
                                key={t}
                                type="button"
                                onClick={() => toggleEditTier(t)}
                                style={{
                                  fontSize: 11.5,
                                  padding: "4px 10px",
                                  borderRadius: 999,
                                  border: `1px solid ${editAllowedTiers.includes(t) ? COLORS.seal : COLORS.grid}`,
                                  background: editAllowedTiers.includes(t) ? "rgba(49,112,142,0.08)" : "transparent",
                                  color: editAllowedTiers.includes(t) ? COLORS.seal : COLORS.inkSoft,
                                  cursor: "pointer",
                                }}
                              >
                                {editAllowedTiers.includes(t) ? "✓ " : ""}
                                {t}
                              </button>
                            ))}
                          </div>

                          <div style={{ fontSize: 11, color: COLORS.inkSoft, marginBottom: 6 }}>
                            {t("admin_course_grants_label", meaningDisplay)}
                          </div>
                          {courseGrants.length > 0 && (
                            <div style={{ display: "flex", flexWrap: "wrap", gap: 6, marginBottom: 8 }}>
                              {courseGrants.map((c) => (
                                <span
                                  key={c}
                                  style={{
                                    display: "inline-flex",
                                    alignItems: "center",
                                    gap: 4,
                                    fontSize: 11.5,
                                    padding: "3px 6px 3px 10px",
                                    borderRadius: 999,
                                    border: `1px solid ${COLORS.seal}`,
                                    background: "rgba(49,112,142,0.08)",
                                    color: COLORS.seal,
                                  }}
                                >
                                  {c}
                                  <button
                                    type="button"
                                    onClick={() => removeCourseGrant(name, c)}
                                    style={{ background: "none", border: "none", color: COLORS.seal, cursor: "pointer", fontSize: 11.5, lineHeight: 1, padding: 0 }}
                                  >
                                    ✕
                                  </button>
                                </span>
                              ))}
                            </div>
                          )}
                          <div style={{ display: "flex", gap: 6, marginBottom: 12 }}>
                            <input
                              value={newCourseGrant}
                              onChange={(e) => setNewCourseGrant(e.target.value)}
                              placeholder={t("admin_course_name_placeholder", meaningDisplay)}
                              style={{ ...inputStyle, fontSize: 12.5, padding: "6px 10px", width: 200 }}
                            />
                            <button
                              type="button"
                              onClick={() => addCourseGrant(name)}
                              className="ghost-btn"
                              style={{ ...ghostBtnStyle, padding: "6px 12px", fontSize: 12 }}
                            >
                              {t("admin_add", meaningDisplay)}
                            </button>
                          </div>
                        </>
                      )}

                      <div style={{ display: "flex", gap: 6 }}>
                        <button
                          type="button"
                          onClick={() => saveListSetting(name)}
                          className="seal-btn"
                          style={{ ...sealBtnStyle, padding: "6px 14px", fontSize: 12 }}
                        >
                          {t("admin_save", meaningDisplay)}
                        </button>
                        <button
                          type="button"
                          onClick={() => setEditingListName(null)}
                          className="ghost-btn"
                          style={{ ...ghostBtnStyle, padding: "6px 14px", fontSize: 12 }}
                        >
                          {t("admin_cancel", meaningDisplay)}
                        </button>
                      </div>
                    </div>
                  ) : (
                    <div style={{ display: "flex", flexWrap: "wrap", alignItems: "center", gap: 10 }}>
                      <div style={{ flex: "1 1 160px", fontSize: 13, color: COLORS.ink, fontWeight: 600 }}>{name}</div>
                      <div style={{ fontSize: 12, color: COLORS.inkSoft }}>
                        {isOpen
                          ? t("admin_open_to_all", meaningDisplay)
                          : setting.admin_only
                          ? t("admin_admin_only_badge", meaningDisplay)
                          : [
                              ...(setting.allowed_tiers || []),
                              ...(courseGrants.length > 0 ? [t("admin_course_count", meaningDisplay, courseGrants.length)] : []),
                            ].join(", ") || t("admin_no_one_can_view", meaningDisplay)}
                      </div>
                      <button
                        type="button"
                        onClick={() => startEditList(name)}
                        className="ghost-btn"
                        style={{ ...ghostBtnStyle, padding: "5px 10px", fontSize: 11.5 }}
                      >
                        {t("admin_edit", meaningDisplay)}
                      </button>
                    </div>
                  )}
                </div>
              );
            })}
            {allListNamesInUse.length === 0 && (
              <div style={{ textAlign: "center", color: COLORS.inkSoft, padding: 20 }}>{t("admin_no_lists_yet", meaningDisplay)}</div>
            )}
          </div>
        )}
      </div>
      )}

      {adminSection === "feedback" && (
      <div style={{ marginTop: 28 }}>
        <div style={{ fontSize: 12.5, fontWeight: 600, color: COLORS.gold, marginBottom: 12, textTransform: "uppercase", letterSpacing: 0.8, textAlign: "center" }}>
          {t("admin_feedback_title", meaningDisplay)}
        </div>
        {feedbackLoading ? (
          <div style={{ textAlign: "center", color: COLORS.inkSoft, padding: 20 }}>{t("loading", meaningDisplay)}</div>
        ) : feedbackList.length === 0 ? (
          <div style={{ textAlign: "center", color: COLORS.inkSoft, padding: 20 }}>{t("admin_feedback_none", meaningDisplay)}</div>
        ) : (
          <div style={{ display: "flex", flexDirection: "column", gap: 10 }}>
            {feedbackList.map((f) => (
              <div key={f.id} style={{ background: COLORS.card, border: `1px solid ${COLORS.hairline}`, borderLeft: `3px solid ${f.read ? COLORS.metadata : COLORS.error}`, borderRadius: 11, padding: "14px 16px" }}>
                <div style={{ display: "flex", justifyContent: "space-between", alignItems: "baseline", gap: 10, marginBottom: 4, flexWrap: "wrap" }}>
                  <span style={{ fontSize: 12, fontWeight: 600, color: COLORS.seal }}>
                    {f.accountEmail
                      ? `${t("admin_feedback_account", meaningDisplay)} ${f.accountEmail}`
                      : t("admin_feedback_guest", meaningDisplay)}
                  </span>
                  <span style={{ fontSize: 11, color: COLORS.metadata }}>{new Date(f.created_at).toLocaleString()}</span>
                </div>
                {f.email && f.email !== f.accountEmail && (
                  <div style={{ fontSize: 11.5, color: COLORS.metadata, marginBottom: 6 }}>
                    {t("admin_feedback_submitted_email", meaningDisplay)} {f.email}
                  </div>
                )}
                {!f.email && !f.accountEmail && (
                  <div style={{ fontSize: 11.5, color: COLORS.metadata, marginBottom: 6 }}>{t("admin_feedback_no_email", meaningDisplay)}</div>
                )}
                <div style={{ fontSize: 13.5, color: COLORS.ink, lineHeight: 1.5, marginBottom: 8, whiteSpace: "pre-wrap" }}>{f.message}</div>
                <div style={{ display: "flex", gap: 6, flexWrap: "wrap" }}>
                  {!f.read && (
                    <button
                      type="button"
                      onClick={() => markFeedbackRead(f.id)}
                      className="ghost-btn"
                      style={{ ...ghostBtnStyle, padding: "4px 10px", fontSize: 11.5, borderColor: COLORS.seal, color: COLORS.seal }}
                    >
                      {t("admin_feedback_mark_read", meaningDisplay)}
                    </button>
                  )}
                  <button
                    type="button"
                    onClick={() => handleDeleteFeedback(f.id)}
                    className="ghost-btn"
                    style={{ ...ghostBtnStyle, padding: "4px 10px", fontSize: 11.5, borderColor: COLORS.error, color: COLORS.error }}
                  >
                    {t("admin_feedback_delete", meaningDisplay)}
                  </button>
                </div>
              </div>
            ))}
          </div>
        )}
      </div>
      )}

      {adminSection === "decks" && (
      <div style={{ marginTop: 28 }}>
        <div style={{ display: "flex", justifyContent: "space-between", alignItems: "center", marginBottom: 12 }}>
          <div style={{ fontSize: 12.5, fontWeight: 600, color: COLORS.gold, textTransform: "uppercase", letterSpacing: 0.8 }}>
            {t("admin_deck_title", meaningDisplay)}
          </div>
          <button type="button" onClick={startNewDeck} className="seal-btn" style={{ ...sealBtnStyle, padding: "6px 14px", fontSize: 12 }}>
            {t("admin_deck_new", meaningDisplay)}
          </button>
        </div>

        {editingDeckId && (
          <div style={{ background: COLORS.card, border: `1px solid ${COLORS.hairline}`, borderRadius: 11, padding: "14px 16px", marginBottom: 14 }}>
            <input
              value={deckName}
              onChange={(e) => setDeckName(e.target.value)}
              placeholder={t("admin_deck_name_placeholder", meaningDisplay)}
              style={{ ...inputStyle, width: "100%", boxSizing: "border-box", marginBottom: 8, fontWeight: 600 }}
            />
            <input
              value={deckDescription}
              onChange={(e) => setDeckDescription(e.target.value)}
              placeholder={t("admin_deck_description_placeholder", meaningDisplay)}
              style={{ ...inputStyle, width: "100%", boxSizing: "border-box", marginBottom: 10 }}
            />

            <div style={{ fontSize: 11.5, color: COLORS.inkSoft, marginBottom: 6 }}>{t("admin_deck_included_lists", meaningDisplay)}</div>
            {deckLists.length === 0 ? (
              <div style={{ fontSize: 12, color: COLORS.metadata, marginBottom: 10 }}>{t("admin_deck_none_included", meaningDisplay)}</div>
            ) : (
              <div style={{ display: "flex", flexWrap: "wrap", gap: 6, marginBottom: 10 }}>
                {deckLists.map((l) => (
                  <span
                    key={`${l.content_type}:${l.list_name}`}
                    style={{ display: "inline-flex", alignItems: "center", gap: 5, background: COLORS.chipBg, border: `1px solid ${COLORS.hairline}`, borderRadius: 999, padding: "3px 9px", fontSize: 11.5 }}
                  >
                    <span style={{ color: COLORS.seal, fontWeight: 600 }}>
                      {l.content_type === "char" ? t("admin_deck_type_char", meaningDisplay) : l.content_type === "word" ? t("admin_deck_type_word", meaningDisplay) : t("admin_deck_type_bushou", meaningDisplay)}
                    </span>
                    · {l.content_type === "bushou" ? displayListName(l.list_name, meaningDisplay) : l.list_name}
                    <button
                      type="button"
                      onClick={() => removeDeckList(l.content_type, l.list_name)}
                      style={{ background: "none", border: "none", color: COLORS.error, cursor: "pointer", padding: 0, fontSize: 12, lineHeight: 1 }}
                    >
                      ✕
                    </button>
                  </span>
                ))}
              </div>
            )}

            <div style={{ display: "flex", gap: 6, marginBottom: 12, flexWrap: "wrap" }}>
              <select value={deckAddType} onChange={(e) => { setDeckAddType(e.target.value); setDeckAddList(""); }} style={{ ...selectStyle, width: 130 }}>
                <option value="char" style={{ background: COLORS.chipBg }}>{t("admin_deck_type_char", meaningDisplay)}</option>
                <option value="word" style={{ background: COLORS.chipBg }}>{t("admin_deck_type_word", meaningDisplay)}</option>
                <option value="bushou" style={{ background: COLORS.chipBg }}>{t("admin_deck_type_bushou", meaningDisplay)}</option>
              </select>
              <select value={deckAddList} onChange={(e) => setDeckAddList(e.target.value)} style={{ ...selectStyle, width: 180 }}>
                <option value="" style={{ background: COLORS.chipBg }}>{t("admin_deck_choose_list", meaningDisplay)}</option>
                {(availableListsByType[deckAddType] || []).map((l) => (
                  <option key={l} value={l} style={{ background: COLORS.chipBg }}>
                    {deckAddType === "bushou" ? displayListName(l, meaningDisplay) : l}
                  </option>
                ))}
              </select>
              <button type="button" onClick={addDeckList} className="ghost-btn" style={{ ...ghostBtnStyle, padding: "6px 12px", fontSize: 12 }}>
                {t("admin_deck_add", meaningDisplay)}
              </button>
            </div>

            {deckMessage && (
              <div style={{ fontSize: 12, fontWeight: 600, color: deckMessage.type === "error" ? COLORS.error : COLORS.seal, marginBottom: 10 }}>
                {deckMessage.text}
              </div>
            )}
            <div style={{ display: "flex", gap: 8 }}>
              <button type="button" onClick={saveDeck} className="seal-btn" style={{ ...sealBtnStyle, padding: "8px 16px", fontSize: 13 }}>
                {t("admin_deck_save", meaningDisplay)}
              </button>
              <button type="button" onClick={() => setEditingDeckId(null)} className="ghost-btn" style={{ ...ghostBtnStyle, padding: "8px 16px", fontSize: 13 }}>
                {t("admin_deck_cancel", meaningDisplay)}
              </button>
            </div>
          </div>
        )}

        {(decks || []).length === 0 ? (
          <div style={{ textAlign: "center", color: COLORS.inkSoft, padding: 20 }}>{t("admin_deck_none", meaningDisplay)}</div>
        ) : (
          <div style={{ display: "flex", flexDirection: "column", gap: 8 }}>
            {(decks || []).map((d) => (
              <div key={d.id} style={{ background: COLORS.card, border: `1px solid ${COLORS.hairline}`, borderRadius: 11, padding: "12px 16px", display: "flex", justifyContent: "space-between", alignItems: "center", gap: 10, flexWrap: "wrap" }}>
                <div>
                  <div style={{ fontSize: 13.5, fontWeight: 600, color: COLORS.ink }}>📦 {d.name}</div>
                  <div style={{ fontSize: 11, color: COLORS.metadata }}>{t("admin_deck_list_count", meaningDisplay, (d.lists || []).length)}</div>
                </div>
                <div style={{ display: "flex", gap: 6 }}>
                  <button type="button" onClick={() => startEditDeck(d)} className="ghost-btn" style={{ ...ghostBtnStyle, padding: "5px 10px", fontSize: 11.5 }}>
                    {t("admin_deck_edit", meaningDisplay)}
                  </button>
                  <button
                    type="button"
                    onClick={() => deleteDeck(d.id)}
                    className="ghost-btn"
                    style={{ ...ghostBtnStyle, padding: "5px 10px", fontSize: 11.5, borderColor: COLORS.error, color: COLORS.error }}
                  >
                    {t("admin_deck_delete", meaningDisplay)}
                  </button>
                </div>
              </div>
            ))}
          </div>
        )}
      </div>
      )}

      {adminSection === "userlib" && (
      <div style={{ marginTop: 28 }}>
        <div style={{ fontSize: 12.5, fontWeight: 600, color: COLORS.gold, marginBottom: 12, textTransform: "uppercase", letterSpacing: 0.8, textAlign: "center" }}>
          {t("admin_nav_userlib", meaningDisplay)}
        </div>

        {!libSelectedUser ? (
          <div style={{ maxWidth: 360, margin: "0 auto" }}>
            <input
              value={libSearchQuery}
              onChange={(e) => setLibSearchQuery(e.target.value)}
              placeholder={t("admin_lib_search_placeholder", meaningDisplay)}
              style={{ ...inputStyle, width: "100%", boxSizing: "border-box" }}
            />
            {libSearchResults.length > 0 && (
              <div style={{ marginTop: 8, display: "flex", flexDirection: "column", gap: 4 }}>
                {libSearchResults.map((u) => (
                  <button
                    key={u.user_id}
                    type="button"
                    onClick={() => selectLibUser(u)}
                    className="ghost-btn"
                    style={{ ...ghostBtnStyle, textAlign: "left", padding: "8px 12px", fontSize: 13 }}
                  >
                    {u.email}
                  </button>
                ))}
              </div>
            )}
          </div>
        ) : (
          <div>
            <div
              style={{
                display: "flex",
                justifyContent: "space-between",
                alignItems: "center",
                gap: 10,
                flexWrap: "wrap",
                background: "rgba(184,134,11,0.12)",
                border: `1.5px solid ${COLORS.gold}`,
                borderRadius: 10,
                padding: "10px 16px",
                marginBottom: 18,
              }}
            >
              <span style={{ fontSize: 13, fontWeight: 700, color: COLORS.gold }}>
                {t("admin_lib_viewing_banner", meaningDisplay, libSelectedUser.email)}
              </span>
              <button type="button" onClick={exitLibView} className="ghost-btn" style={{ ...ghostBtnStyle, padding: "5px 12px", fontSize: 12 }}>
                {t("admin_lib_exit", meaningDisplay)}
              </button>
            </div>

            {libCopyMessage && (
              <div style={{ textAlign: "center", fontSize: 13, fontWeight: 600, color: libCopyMessage.type === "error" ? COLORS.error : COLORS.seal, marginBottom: 16 }}>
                {libCopyMessage.text}
              </div>
            )}

            <div style={{ textAlign: "center", marginBottom: 16 }}>
              <label style={{ display: "inline-flex", alignItems: "center", gap: 6, fontSize: 12.5, color: COLORS.ink, cursor: "pointer" }}>
                <input type="checkbox" checked={libOnlyNew} onChange={(e) => setLibOnlyNew(e.target.checked)} />
                {t("admin_lib_only_new", meaningDisplay)}
              </label>
            </div>

            {libLoading ? (
              <div style={{ textAlign: "center", color: COLORS.inkSoft, padding: 20 }}>{t("loading", meaningDisplay)}</div>
            ) : (
              <div style={{ maxWidth: 560, margin: "0 auto" }}>
                {[
                  { type: "char", title: t("admin_lib_chars_title", meaningDisplay), data: libUserChars },
                  { type: "word", title: t("admin_lib_words_title", meaningDisplay), data: libUserWords },
                  { type: "bushou", title: t("admin_lib_bushou_title", meaningDisplay), data: libUserBushou },
                ].map(({ type, title, data }) => {
                  const listNames = Array.from(new Set((data || []).flatMap((item) => item.lists || []))).sort((a, b) => {
                    if (type === "bushou") {
                      const numA = parseInt(a, 10);
                      const numB = parseInt(b, 10);
                      if (!isNaN(numA) && !isNaN(numB)) return numA - numB;
                    }
                    return a.localeCompare(b, "vi");
                  });
                  const unlisted = (data || []).filter((item) => (!item.lists || item.lists.length === 0) && (!libOnlyNew || isNewToMe(type, item)));
                  return (
                    <div key={type} style={{ marginBottom: 26 }}>
                      <div style={{ fontSize: 12, fontWeight: 700, color: COLORS.ink, textTransform: "uppercase", letterSpacing: 0.5, marginBottom: 10 }}>
                        {title}
                      </div>
                      {(!data || data.length === 0) ? (
                        <div style={{ fontSize: 12.5, color: COLORS.inkSoft }}>{t("admin_lib_no_content", meaningDisplay)}</div>
                      ) : (
                        <>
                          {listNames.map((listName) => {
                            const itemsInList = data.filter((item) => (item.lists || []).includes(listName) && (!libOnlyNew || isNewToMe(type, item)));
                            if (itemsInList.length === 0) return null;
                            return (
                              <div key={listName} style={{ marginBottom: 14 }}>
                                <div style={{ display: "flex", justifyContent: "space-between", alignItems: "center", marginBottom: 6 }}>
                                  <span style={{ fontSize: 12, fontWeight: 600, color: COLORS.seal }}>
                                    {t("admin_lib_lists_in", meaningDisplay)} {type === "bushou" ? displayListName(listName, meaningDisplay) : listName}
                                  </span>
                                  <button
                                    type="button"
                                    onClick={() => copyList(type, listName)}
                                    style={{
                                      fontSize: 11,
                                      fontWeight: 600,
                                      color: COLORS.seal,
                                      background: COLORS.chipBg,
                                      border: `1px solid ${COLORS.hairline}`,
                                      borderRadius: 999,
                                      padding: "3px 10px",
                                      cursor: "pointer",
                                    }}
                                  >
                                    {t("admin_lib_copy_list", meaningDisplay)}
                                  </button>
                                </div>
                                <div style={{ display: "flex", flexDirection: "column", gap: 6 }}>
                                  {itemsInList.map((item) => {
                                    const isNew = isNewToMe(type, item);
                                    return (
                                    <div
                                      key={item.char || item.word}
                                      style={{ display: "flex", justifyContent: "space-between", alignItems: "center", gap: 10, background: COLORS.card, border: `1px solid ${isNew ? COLORS.seal : COLORS.hairline}`, borderRadius: 8, padding: "8px 12px" }}
                                    >
                                      <div style={{ fontSize: 12.5, color: COLORS.ink, display: "flex", alignItems: "center", gap: 6 }}>
                                        {isNew && (
                                          <span style={{ fontSize: 9.5, fontWeight: 700, color: "#FBF9EF", background: COLORS.seal, borderRadius: 4, padding: "1px 5px" }}>
                                            {t("admin_lib_new_badge", meaningDisplay)}
                                          </span>
                                        )}
                                        <span style={{ fontFamily: "'Noto Serif SC', 'STKaiti', 'Kaiti SC', serif", fontWeight: 700 }}>
                                          {item.char || item.word}
                                        </span>
                                        {item.pinyin} · {item.meaning}
                                      </div>
                                      <button
                                        type="button"
                                        onClick={() => copyItem(type, item)}
                                        className="ghost-btn"
                                        style={{ ...ghostBtnStyle, padding: "3px 10px", fontSize: 11 }}
                                      >
                                        {t("admin_lib_copy", meaningDisplay)}
                                      </button>
                                    </div>
                                    );
                                  })}
                                </div>
                              </div>
                            );
                          })}

                          {unlisted.length > 0 && (
                            <div style={{ marginBottom: 14 }}>
                              <div style={{ fontSize: 12, fontWeight: 600, color: COLORS.metadata, marginBottom: 6 }}>
                                {t("admin_lib_unlisted", meaningDisplay)}
                              </div>
                              <div style={{ display: "flex", flexDirection: "column", gap: 6 }}>
                                {unlisted.map((item) => {
                                  const isNew = isNewToMe(type, item);
                                  return (
                                  <div
                                    key={item.char || item.word}
                                    style={{ display: "flex", justifyContent: "space-between", alignItems: "center", gap: 10, background: COLORS.card, border: `1px solid ${isNew ? COLORS.seal : COLORS.hairline}`, borderRadius: 8, padding: "8px 12px" }}
                                  >
                                    <div style={{ fontSize: 12.5, color: COLORS.ink, display: "flex", alignItems: "center", gap: 6 }}>
                                      {isNew && (
                                        <span style={{ fontSize: 9.5, fontWeight: 700, color: "#FBF9EF", background: COLORS.seal, borderRadius: 4, padding: "1px 5px" }}>
                                          {t("admin_lib_new_badge", meaningDisplay)}
                                        </span>
                                      )}
                                      <span style={{ fontFamily: "'Noto Serif SC', 'STKaiti', 'Kaiti SC', serif", fontWeight: 700 }}>
                                        {item.char || item.word}
                                      </span>
                                      {item.pinyin} · {item.meaning}
                                    </div>
                                    <button
                                      type="button"
                                      onClick={() => copyItem(type, item)}
                                      className="ghost-btn"
                                      style={{ ...ghostBtnStyle, padding: "3px 10px", fontSize: 11 }}
                                    >
                                      {t("admin_lib_copy", meaningDisplay)}
                                    </button>
                                  </div>
                                  );
                                })}
                              </div>
                            </div>
                          )}
                        </>
                      )}
                    </div>
                  );
                })}

                <div>
                  <div style={{ fontSize: 12, fontWeight: 700, color: COLORS.ink, textTransform: "uppercase", letterSpacing: 0.5, marginBottom: 8 }}>
                    {t("admin_lib_decks_title", meaningDisplay)}
                  </div>
                  {(!libUserDecks || libUserDecks.length === 0) ? (
                    <div style={{ fontSize: 12.5, color: COLORS.inkSoft }}>{t("admin_lib_no_content", meaningDisplay)}</div>
                  ) : (
                    <div style={{ display: "flex", flexDirection: "column", gap: 6 }}>
                      {libUserDecks.map((d) => (
                        <div key={d.id} style={{ display: "flex", justifyContent: "space-between", alignItems: "center", gap: 10, background: COLORS.card, border: `1px solid ${COLORS.hairline}`, borderRadius: 8, padding: "8px 12px" }}>
                          <div style={{ fontSize: 12.5, color: COLORS.ink }}>
                            📦 {d.name} <span style={{ color: COLORS.metadata }}>({t("admin_deck_list_count", meaningDisplay, (d.lists || []).length)})</span>
                          </div>
                          <button type="button" onClick={() => copyDeck(d)} className="ghost-btn" style={{ ...ghostBtnStyle, padding: "3px 10px", fontSize: 11 }}>
                            {t("admin_lib_copy_deck", meaningDisplay)}
                          </button>
                        </div>
                      ))}
                    </div>
                  )}
                </div>
              </div>
            )}
          </div>
        )}
      </div>
      )}

      {adminSection === "suggestions" && (
      <div style={{ marginTop: 28 }}>
        <div style={{ fontSize: 12.5, fontWeight: 600, color: COLORS.gold, marginBottom: 12, textTransform: "uppercase", letterSpacing: 0.8, textAlign: "center" }}>
          {t("admin_suggestions_title", meaningDisplay)}
        </div>

        <div style={{ display: "flex", justifyContent: "center", gap: 6, marginBottom: 16, flexWrap: "wrap" }}>
          {["all", "new", "read", "revised", "ignored"].map((s) => (
            <button
              key={s}
              type="button"
              onClick={() => setSuggestionStatusFilter(s)}
              style={{
                padding: "5px 12px",
                borderRadius: 999,
                border: `1.5px solid ${suggestionStatusFilter === s ? COLORS.seal : COLORS.hairline}`,
                background: suggestionStatusFilter === s ? COLORS.seal : "transparent",
                color: suggestionStatusFilter === s ? "#FBF9EF" : COLORS.inkSoft,
                fontSize: 12,
                fontWeight: 600,
                cursor: "pointer",
              }}
            >
              {s === "all"
                ? t("admin_suggestion_filter_all", meaningDisplay)
                : s === "new"
                ? t("admin_suggestion_status_new", meaningDisplay)
                : s === "read"
                ? t("admin_suggestion_status_read", meaningDisplay)
                : s === "revised"
                ? t("admin_suggestion_status_revised", meaningDisplay)
                : t("admin_suggestion_status_ignored", meaningDisplay)}
            </button>
          ))}
        </div>

        {suggestionsLoading ? (
          <div style={{ textAlign: "center", color: COLORS.inkSoft, padding: 20 }}>{t("loading", meaningDisplay)}</div>
        ) : (
          <div style={{ display: "flex", flexDirection: "column", gap: 10 }}>
            {suggestions
              .filter((s) => suggestionStatusFilter === "all" || s.status === suggestionStatusFilter)
              .map((s) => {
                const item = findSuggestionItem(s);
                const typeLabel =
                  s.content_type === "char"
                    ? t("admin_suggestion_type_char", meaningDisplay)
                    : s.content_type === "word"
                    ? t("admin_suggestion_type_word", meaningDisplay)
                    : t("admin_suggestion_type_bushou", meaningDisplay);
                const statusColor =
                  s.status === "new" ? COLORS.error : s.status === "read" ? COLORS.gold : s.status === "revised" ? COLORS.seal : COLORS.metadata;
                const statusLabel =
                  s.status === "new"
                    ? t("admin_suggestion_status_new", meaningDisplay)
                    : s.status === "read"
                    ? t("admin_suggestion_status_read", meaningDisplay)
                    : s.status === "revised"
                    ? t("admin_suggestion_status_revised", meaningDisplay)
                    : t("admin_suggestion_status_ignored", meaningDisplay);
                return (
                  <div key={s.id} style={{ background: COLORS.card, border: `1px solid ${COLORS.hairline}`, borderLeft: `3px solid ${statusColor}`, borderRadius: 11, padding: "14px 16px" }}>
                    <div style={{ display: "flex", justifyContent: "space-between", alignItems: "baseline", gap: 10, marginBottom: 6, flexWrap: "wrap" }}>
                      <span style={{ fontSize: 12.5, fontWeight: 700, color: COLORS.ink }}>
                        <span style={{ color: COLORS.seal }}>{typeLabel}</span>
                        {" · "}
                        {item ? (
                          <>
                            <span style={{ fontFamily: "'Noto Serif SC', 'STKaiti', 'Kaiti SC', serif" }}>{item.char || item.word}</span>
                            {" ("}{item.pinyin}{" · "}{item.meaning}{")"}
                          </>
                        ) : (
                          <>
                            <span style={{ fontFamily: "'Noto Serif SC', 'STKaiti', 'Kaiti SC', serif" }}>{s.item_key}</span>
                            {" "}{t("admin_suggestion_unknown_item", meaningDisplay)}
                          </>
                        )}
                      </span>
                      <span style={{ fontSize: 10.5, fontWeight: 700, color: statusColor, textTransform: "uppercase" }}>{statusLabel}</span>
                    </div>
                    <div style={{ fontSize: 13.5, color: COLORS.ink, lineHeight: 1.5, marginBottom: 6, whiteSpace: "pre-wrap" }}>{s.message}</div>
                    <div style={{ fontSize: 11, color: COLORS.metadata, marginBottom: 10 }}>
                      {s.email || t("admin_suggestion_no_email", meaningDisplay)} · {new Date(s.created_at).toLocaleString()}
                    </div>
                    <div style={{ display: "flex", gap: 6, flexWrap: "wrap" }}>
                      <button type="button" onClick={() => setSuggestionStatus(s.id, "read")} className="ghost-btn" style={{ ...ghostBtnStyle, padding: "4px 10px", fontSize: 11.5 }}>
                        {t("admin_suggestion_mark_read", meaningDisplay)}
                      </button>
                      <button type="button" onClick={() => setSuggestionStatus(s.id, "revised")} className="ghost-btn" style={{ ...ghostBtnStyle, padding: "4px 10px", fontSize: 11.5, borderColor: COLORS.seal, color: COLORS.seal }}>
                        {t("admin_suggestion_mark_revised", meaningDisplay)}
                      </button>
                      <button type="button" onClick={() => setSuggestionStatus(s.id, "ignored")} className="ghost-btn" style={{ ...ghostBtnStyle, padding: "4px 10px", fontSize: 11.5 }}>
                        {t("admin_suggestion_mark_ignored", meaningDisplay)}
                      </button>
                      <button
                        type="button"
                        onClick={() => deleteSuggestion(s.id)}
                        className="ghost-btn"
                        style={{ ...ghostBtnStyle, padding: "4px 10px", fontSize: 11.5, borderColor: COLORS.error, color: COLORS.error }}
                      >
                        {t("admin_blog_delete", meaningDisplay)}
                      </button>
                    </div>
                  </div>
                );
              })}
            {suggestions.filter((s) => suggestionStatusFilter === "all" || s.status === suggestionStatusFilter).length === 0 && (
              <div style={{ textAlign: "center", color: COLORS.inkSoft, padding: 20 }}>{t("admin_suggestion_none", meaningDisplay)}</div>
            )}
          </div>
        )}
      </div>
      )}

      {adminSection === "blog" && (
      <div style={{ marginTop: 28 }}>
        <div style={{ display: "flex", justifyContent: "space-between", alignItems: "center", marginBottom: 12 }}>
          <div style={{ fontSize: 12.5, fontWeight: 600, color: COLORS.gold, textTransform: "uppercase", letterSpacing: 0.8 }}>
            {t("admin_blog_title", meaningDisplay)}
          </div>
          <button type="button" onClick={startNewBlogPost} className="seal-btn" style={{ ...sealBtnStyle, padding: "6px 14px", fontSize: 12 }}>
            {t("admin_blog_new_post", meaningDisplay)}
          </button>
        </div>

        {editingBlogId && (
          <div style={{ background: COLORS.card, border: `1px solid ${COLORS.hairline}`, borderRadius: 11, padding: "14px 16px", marginBottom: 14 }}>
            <input
              value={blogTitle}
              onChange={(e) => setBlogTitle(e.target.value)}
              placeholder={t("admin_blog_title_placeholder", meaningDisplay)}
              style={{ ...inputStyle, width: "100%", boxSizing: "border-box", marginBottom: 8, fontWeight: 600 }}
            />
            <div style={{ display: "flex", gap: 8, marginBottom: 8, flexWrap: "wrap" }}>
              <select value={blogCategory} onChange={(e) => setBlogCategory(e.target.value)} style={{ ...selectStyle, width: 170 }}>
                <option value="news" style={{ background: COLORS.chipBg }}>{t("blog_cat_news", meaningDisplay)}</option>
                <option value="resources" style={{ background: COLORS.chipBg }}>{t("blog_cat_resources", meaningDisplay)}</option>
                <option value="founder" style={{ background: COLORS.chipBg }}>{t("blog_cat_founder", meaningDisplay)}</option>
              </select>
              <label style={{ display: "flex", alignItems: "center", gap: 6, fontSize: 12.5, color: COLORS.ink, cursor: "pointer" }}>
                <input type="checkbox" checked={blogPublished} onChange={(e) => setBlogPublished(e.target.checked)} />
                {blogPublished ? t("admin_blog_published", meaningDisplay) : t("admin_blog_draft", meaningDisplay)}
              </label>
            </div>
            <div style={{ marginBottom: 8 }}>
              <RichTextEditor
                value={blogBody}
                onChange={setBlogBody}
                placeholder={t("admin_blog_body_placeholder", meaningDisplay)}
              />
            </div>
            <input
              value={blogLink}
              onChange={(e) => setBlogLink(e.target.value)}
              placeholder={t("admin_blog_link_placeholder", meaningDisplay)}
              style={{ ...inputStyle, width: "100%", boxSizing: "border-box", marginBottom: 10 }}
            />
            {blogMessage && (
              <div style={{ fontSize: 12, fontWeight: 600, color: blogMessage.type === "error" ? COLORS.error : COLORS.seal, marginBottom: 10 }}>
                {blogMessage.text}
              </div>
            )}
            <div style={{ display: "flex", gap: 8 }}>
              <button type="button" onClick={saveBlogPost} className="seal-btn" style={{ ...sealBtnStyle, padding: "8px 16px", fontSize: 13 }}>
                {t("admin_blog_save", meaningDisplay)}
              </button>
              <button type="button" onClick={() => setEditingBlogId(null)} className="ghost-btn" style={{ ...ghostBtnStyle, padding: "8px 16px", fontSize: 13 }}>
                {t("admin_blog_cancel", meaningDisplay)}
              </button>
            </div>
          </div>
        )}

        {blogLoading ? (
          <div style={{ textAlign: "center", color: COLORS.inkSoft, padding: 20 }}>{t("loading", meaningDisplay)}</div>
        ) : blogPosts.length === 0 ? (
          <div style={{ textAlign: "center", color: COLORS.inkSoft, padding: 20 }}>{t("admin_blog_none", meaningDisplay)}</div>
        ) : (
          <div style={{ display: "flex", flexDirection: "column", gap: 8 }}>
            {blogPosts.map((post) => (
              <div key={post.id} style={{ background: COLORS.card, border: `1px solid ${COLORS.hairline}`, borderRadius: 11, padding: "12px 16px", display: "flex", justifyContent: "space-between", alignItems: "center", gap: 10, flexWrap: "wrap" }}>
                <div>
                  <div style={{ fontSize: 13.5, fontWeight: 600, color: COLORS.ink }}>{post.title}</div>
                  <div style={{ fontSize: 11, color: post.published ? COLORS.seal : COLORS.metadata }}>
                    {post.published ? t("admin_blog_published", meaningDisplay) : t("admin_blog_draft", meaningDisplay)} · {post.category}
                  </div>
                </div>
                <div style={{ display: "flex", gap: 6 }}>
                  <button
                    type="button"
                    onClick={() => toggleBlogPublish(post)}
                    className="seal-btn"
                    style={{
                      ...sealBtnStyle,
                      padding: "5px 10px",
                      fontSize: 11.5,
                      background: post.published ? COLORS.metadata : COLORS.seal,
                    }}
                  >
                    {post.published ? t("admin_blog_unpublish_action", meaningDisplay) : t("admin_blog_publish_action", meaningDisplay)}
                  </button>
                  <button type="button" onClick={() => startEditBlogPost(post)} className="ghost-btn" style={{ ...ghostBtnStyle, padding: "5px 10px", fontSize: 11.5 }}>
                    {t("admin_blog_edit", meaningDisplay)}
                  </button>
                  <button
                    type="button"
                    onClick={() => deleteBlogPost(post.id)}
                    className="ghost-btn"
                    style={{ ...ghostBtnStyle, padding: "5px 10px", fontSize: 11.5, borderColor: COLORS.error, color: COLORS.error }}
                  >
                    {t("admin_blog_delete", meaningDisplay)}
                  </button>
                </div>
              </div>
            ))}
          </div>
        )}
      </div>
      )}

      {adminSection === "comments" && (
      <div style={{ marginTop: 28 }}>
        <div style={{ fontSize: 12.5, fontWeight: 600, color: COLORS.gold, marginBottom: 12, textTransform: "uppercase", letterSpacing: 0.8, textAlign: "center" }}>
          {t("admin_comments_title", meaningDisplay)}
        </div>

        <div style={{ display: "flex", justifyContent: "center", gap: 6, marginBottom: 16, flexWrap: "wrap" }}>
          {["pending", "approved", "rejected", "all"].map((s) => (
            <button
              key={s}
              type="button"
              onClick={() => setCommentStatusFilter(s)}
              style={{
                padding: "5px 12px",
                borderRadius: 999,
                border: `1.5px solid ${commentStatusFilter === s ? COLORS.seal : COLORS.hairline}`,
                background: commentStatusFilter === s ? COLORS.seal : "transparent",
                color: commentStatusFilter === s ? "#FBF9EF" : COLORS.inkSoft,
                fontSize: 12,
                fontWeight: 600,
                cursor: "pointer",
              }}
            >
              {s === "pending"
                ? t("admin_comment_filter_pending", meaningDisplay)
                : s === "approved"
                ? t("admin_comment_filter_approved", meaningDisplay)
                : s === "rejected"
                ? t("admin_comment_filter_rejected", meaningDisplay)
                : t("admin_comment_filter_all", meaningDisplay)}
            </button>
          ))}
        </div>

        {commentsLoading ? (
          <div style={{ textAlign: "center", color: COLORS.inkSoft, padding: 20 }}>{t("loading", meaningDisplay)}</div>
        ) : (
          <div style={{ display: "flex", flexDirection: "column", gap: 10 }}>
            {comments
              .filter((c) => commentStatusFilter === "all" || c.status === commentStatusFilter)
              .map((c) => {
                const statusColor = c.status === "pending" ? COLORS.gold : c.status === "approved" ? COLORS.seal : COLORS.metadata;
                return (
                  <div key={c.id} style={{ background: COLORS.card, border: `1px solid ${COLORS.hairline}`, borderLeft: `3px solid ${statusColor}`, borderRadius: 11, padding: "14px 16px" }}>
                    <div style={{ fontSize: 11.5, color: COLORS.metadata, marginBottom: 6 }}>
                      {t("admin_comment_on_post", meaningDisplay)} {c.postTitle || t("admin_comment_unknown_post", meaningDisplay)}
                    </div>
                    <div style={{ display: "flex", justifyContent: "space-between", alignItems: "baseline", gap: 10, marginBottom: 6, flexWrap: "wrap" }}>
                      <span style={{ fontSize: 13, fontWeight: 700, color: COLORS.ink }}>{c.name}</span>
                      <span style={{ fontSize: 10.5, color: COLORS.metadata }}>{new Date(c.created_at).toLocaleString()}</span>
                    </div>
                    <div style={{ fontSize: 13.5, color: COLORS.ink, lineHeight: 1.5, marginBottom: 6, whiteSpace: "pre-wrap" }}>{c.message}</div>
                    <div style={{ fontSize: 11, color: COLORS.metadata, marginBottom: 10 }}>{c.email || t("admin_comment_no_email", meaningDisplay)}</div>
                    <div style={{ display: "flex", gap: 6, flexWrap: "wrap" }}>
                      {c.status !== "approved" && (
                        <button type="button" onClick={() => setCommentStatus(c.id, "approved")} className="ghost-btn" style={{ ...ghostBtnStyle, padding: "4px 10px", fontSize: 11.5, borderColor: COLORS.seal, color: COLORS.seal }}>
                          {t("admin_comment_approve", meaningDisplay)}
                        </button>
                      )}
                      {c.status !== "rejected" && (
                        <button type="button" onClick={() => setCommentStatus(c.id, "rejected")} className="ghost-btn" style={{ ...ghostBtnStyle, padding: "4px 10px", fontSize: 11.5 }}>
                          {t("admin_comment_reject", meaningDisplay)}
                        </button>
                      )}
                      <button
                        type="button"
                        onClick={() => deleteComment(c.id)}
                        className="ghost-btn"
                        style={{ ...ghostBtnStyle, padding: "4px 10px", fontSize: 11.5, borderColor: COLORS.error, color: COLORS.error }}
                      >
                        {t("admin_comment_delete", meaningDisplay)}
                      </button>
                    </div>
                  </div>
                );
              })}
            {comments.filter((c) => commentStatusFilter === "all" || c.status === commentStatusFilter).length === 0 && (
              <div style={{ textAlign: "center", color: COLORS.inkSoft, padding: 20 }}>{t("admin_comment_none", meaningDisplay)}</div>
            )}
          </div>
        )}
      </div>
      )}

      {adminSection === "messages" && (
      <div style={{ marginTop: 28 }}>
        <div style={{ fontSize: 12.5, fontWeight: 600, color: COLORS.gold, marginBottom: 12, textTransform: "uppercase", letterSpacing: 0.8, textAlign: "center" }}>
          {t("admin_messages_title", meaningDisplay)}
        </div>

        <div style={{ display: "flex", gap: 20, flexWrap: "wrap" }}>
          <div style={{ width: 220, flexShrink: 0 }}>
            <input
              value={msgSearchQuery}
              onChange={(e) => setMsgSearchQuery(e.target.value)}
              placeholder={t("admin_messages_search_placeholder", meaningDisplay)}
              style={{ ...inputStyle, width: "100%", boxSizing: "border-box", fontSize: 12.5, marginBottom: 8 }}
            />
            {msgSearchResults.length > 0 && (
              <div style={{ display: "flex", flexDirection: "column", gap: 4, marginBottom: 10 }}>
                {msgSearchResults.map((u) => (
                  <button
                    key={u.user_id}
                    type="button"
                    onClick={() => selectMsgThread(u)}
                    className="ghost-btn"
                    style={{ ...ghostBtnStyle, textAlign: "left", padding: "6px 10px", fontSize: 12 }}
                  >
                    {t("admin_messages_new_thread", meaningDisplay)}: {u.email}
                  </button>
                ))}
              </div>
            )}

            {msgThreads === null ? (
              <div style={{ fontSize: 12.5, color: COLORS.inkSoft, textAlign: "center", padding: 12 }}>{t("loading", meaningDisplay)}</div>
            ) : msgThreads.length === 0 ? (
              <div style={{ fontSize: 12.5, color: COLORS.inkSoft, textAlign: "center", padding: 12 }}>{t("admin_messages_none", meaningDisplay)}</div>
            ) : (
              <div style={{ display: "flex", flexDirection: "column", gap: 4 }}>
                {msgThreads.map((th) => (
                  <button
                    key={th.user_id}
                    type="button"
                    onClick={() => selectMsgThread(th)}
                    style={{
                      textAlign: "left",
                      background: msgSelectedUser && msgSelectedUser.user_id === th.user_id ? COLORS.chipBg : "none",
                      border: "none",
                      borderLeft: `3px solid ${msgSelectedUser && msgSelectedUser.user_id === th.user_id ? COLORS.seal : "transparent"}`,
                      borderRadius: 6,
                      padding: "8px 10px",
                      cursor: "pointer",
                      position: "relative",
                    }}
                  >
                    <div style={{ fontSize: 12.5, fontWeight: 600, color: COLORS.ink, overflow: "hidden", textOverflow: "ellipsis", whiteSpace: "nowrap" }}>
                      {th.email}
                      {th.unread > 0 && (
                        <span style={{ marginLeft: 6, fontSize: 9.5, fontWeight: 700, color: "#FBF9EF", background: COLORS.error, borderRadius: 999, padding: "1px 6px" }}>
                          {th.unread}
                        </span>
                      )}
                    </div>
                    <div style={{ fontSize: 11, color: COLORS.metadata, overflow: "hidden", textOverflow: "ellipsis", whiteSpace: "nowrap" }}>
                      {th.latest.sender === "admin" ? `${t("msg_from_you", meaningDisplay)}: ` : ""}
                      {th.latest.message}
                    </div>
                  </button>
                ))}
              </div>
            )}
          </div>

          <div style={{ flex: 1, minWidth: 260 }}>
            {!msgSelectedUser ? (
              <div style={{ textAlign: "center", color: COLORS.inkSoft, padding: 30 }}>{t("admin_messages_pick_user", meaningDisplay)}</div>
            ) : (
              <div>
                <div style={{ fontSize: 13, fontWeight: 700, color: COLORS.ink, marginBottom: 10 }}>{msgSelectedUser.email}</div>
                <div
                  style={{
                    border: `1px solid ${COLORS.hairline}`,
                    borderRadius: 12,
                    padding: "14px 16px",
                    minHeight: 240,
                    maxHeight: 420,
                    overflowY: "auto",
                    marginBottom: 12,
                    display: "flex",
                    flexDirection: "column",
                    gap: 10,
                  }}
                >
                  {msgThread === null ? (
                    <div style={{ textAlign: "center", color: COLORS.inkSoft, fontSize: 13 }}>{t("loading", meaningDisplay)}</div>
                  ) : msgThread.length === 0 ? (
                    <div style={{ textAlign: "center", color: COLORS.inkSoft, fontSize: 13 }}>{t("msg_empty_state", meaningDisplay)}</div>
                  ) : (
                    msgThread.map((m) => (
                      <div key={m.id} style={{ alignSelf: m.sender === "admin" ? "flex-end" : "flex-start", maxWidth: "80%" }}>
                        <div style={{ fontSize: 10, color: COLORS.metadata, marginBottom: 2, textAlign: m.sender === "admin" ? "right" : "left" }}>
                          {m.sender === "admin" ? t("msg_from_you", meaningDisplay) : msgSelectedUser.email} · {new Date(m.created_at).toLocaleString()}
                        </div>
                        <div
                          style={{
                            background: m.sender === "admin" ? COLORS.seal : COLORS.chipBg,
                            color: m.sender === "admin" ? "#FBF9EF" : COLORS.ink,
                            borderRadius: 10,
                            padding: "8px 12px",
                            fontSize: 13,
                            lineHeight: 1.5,
                            whiteSpace: "pre-wrap",
                          }}
                        >
                          {m.message}
                        </div>
                      </div>
                    ))
                  )}
                </div>
                <form onSubmit={sendMsgReply} style={{ display: "flex", gap: 8 }}>
                  <input
                    value={msgDraft}
                    onChange={(e) => setMsgDraft(e.target.value)}
                    placeholder={t("msg_placeholder", meaningDisplay)}
                    style={{ ...inputStyle, flex: 1, boxSizing: "border-box" }}
                  />
                  <button type="submit" className="seal-btn" style={{ ...sealBtnStyle, padding: "8px 18px" }} disabled={msgSending}>
                    {t("msg_send", meaningDisplay)}
                  </button>
                </form>
              </div>
            )}
          </div>
        </div>
      </div>
      )}
      </div>
    </div>
  );
}

function WordListPanel({ wordList, characterList, findBushou, onAddWord, onDeleteWord, onDeleteWordFromOfficial, isAdmin, officialWordKeys, overrideWordKeys, onPromoteWord, onWithdrawWord, checkListAccess, onViewPremium, meaningDisplay, userId }) {
  const [query, setQuery] = useState("");
  const [listFilter, setListFilter] = useState("Tất cả");
  const [defaultFilter, setDefaultFilter] = useState("all"); // all | official | pending
  const [exportMessage, setExportMessage] = useState(null);
  const [lockedListName, setLockedListName] = useState(null);

  const allLists = useMemo(() => {
    const set = new Set();
    (wordList || []).forEach((w) => (w.lists || []).forEach((l) => set.add(l.trim())));
    return Array.from(set).sort((a, b) => a.localeCompare(b, "vi"));
  }, [wordList]);

  const filtered = (wordList || []).filter((w) => {
    if (listFilter !== "Tất cả" && !(w.lists || []).some((l) => l.trim() === listFilter)) return false;
    if (isAdmin && defaultFilter !== "all") {
      const isOfficial = officialWordKeys ? officialWordKeys.has(w.word) : false;
      const hasOverride = overrideWordKeys ? overrideWordKeys.has(w.word) : false;
      const isCleanlyPublished = isOfficial && !hasOverride;
      if (defaultFilter === "official" && !isCleanlyPublished) return false;
      if (defaultFilter === "pending" && isCleanlyPublished) return false;
    }
    const q = query.trim().toLowerCase();
    if (!q) return true;
    return (
      w.word.includes(q) ||
      (w.pinyin || "").toLowerCase().includes(q) ||
      (w.meaning || "").toLowerCase().includes(q) ||
      (w.sv || "").toLowerCase().includes(q)
    );
  });

  function handleExportExcel() {
    setExportMessage(null);
    try {
      if (filtered.length === 0) {
        setExportMessage({ type: "error", text: t("vocab_export_none", meaningDisplay) });
        return;
      }
      const rows = filtered.map((w) => ({
        "Từ vựng": w.word,
        Pinyin: w.pinyin,
        "Nghĩa (English)": w.meaning,
        "Nghĩa (Tiếng Việt)": w.meaning_vi || "",
        "Âm Hán Việt": w.sv,
        "Danh sách": (w.lists || []).join(", "),
        "Chữ cấu thành": (w.chars || []).join(" + "),
      }));
      const worksheet = XLSX.utils.json_to_sheet(rows);
      worksheet["!cols"] = [
        { wch: 12 },
        { wch: 14 },
        { wch: 32 },
        { wch: 32 },
        { wch: 16 },
        { wch: 20 },
        { wch: 20 },
      ];
      const workbook = XLSX.utils.book_new();
      XLSX.utils.book_append_sheet(workbook, worksheet, "Tu vung");
      const dateStr = new Date().toISOString().slice(0, 10);
      const suffix = listFilter !== "Tất cả" ? `-${listFilter}` : "";
      XLSX.writeFile(workbook, `tu-vung${suffix}-${dateStr}.xlsx`);
      setExportMessage({ type: "success", text: t("vocab_export_success", meaningDisplay, filtered.length) });
    } catch (err) {
      console.error("Export to Excel failed:", err);
      setExportMessage({ type: "error", text: t("hanzi_export_fail", meaningDisplay) });
    }
  }

  return (
    <div>
      <div style={{ fontSize: 12.5, fontWeight: 600, color: COLORS.gold, marginBottom: 12, textTransform: "uppercase", letterSpacing: 0.8, textAlign: "center" }}>
        {userId ? t("vocab_panel_title", meaningDisplay) : t("vocab_panel_title_guest", meaningDisplay)}
      </div>

      {(!wordList || wordList.length === 0) ? (
        <div style={{ textAlign: "center", color: COLORS.inkSoft, fontSize: 13, padding: 30 }}>
          {t("vocab_empty", meaningDisplay)}
        </div>
      ) : (
        <>
          <div style={{ display: "flex", justifyContent: "center", gap: 8, flexWrap: "wrap", marginBottom: 10 }}>
            <input
              value={query}
              onChange={(e) => setQuery(e.target.value)}
              placeholder={t("vocab_search_placeholder", meaningDisplay)}
              style={{ ...inputStyle, width: 260, textAlign: "center" }}
            />
            <select
              value={listFilter}
              onChange={(e) => {
                const next = e.target.value;
                if (!isAdmin && next !== "Tất cả" && checkListAccess && !checkListAccess(next)) {
                  setLockedListName(next);
                  return;
                }
                setListFilter(next);
              }}
              style={{ ...selectStyle, width: 190, flex: "none" }}
            >
              <option value="Tất cả" style={{ background: COLORS.chipBg, color: COLORS.ink, fontWeight: 700 }}>{t("play_all_lists", meaningDisplay)}</option>
              {allLists.map((l) => (
                <option key={l} value={l} style={{ background: COLORS.chipBg, color: COLORS.ink, fontWeight: 700 }}>
                  {!isAdmin && checkListAccess && !checkListAccess(l) ? `🔒 ${l}` : displayListName(l, meaningDisplay)}
                </option>
              ))}
            </select>
            {lockedListName && (
              <ListLockedModal listName={lockedListName} onClose={() => setLockedListName(null)} onViewPremium={onViewPremium} />
            )}
            {isAdmin && (
              <select value={defaultFilter} onChange={(e) => setDefaultFilter(e.target.value)} style={{ ...selectStyle, width: 190, flex: "none" }}>
                <option value="all" style={{ background: COLORS.chipBg, color: COLORS.ink, fontWeight: 700 }}>{t("hanzi_all_status", meaningDisplay)}</option>
                <option value="official" style={{ background: COLORS.chipBg, color: COLORS.ink, fontWeight: 700 }}>{t("hanzi_is_official", meaningDisplay)}</option>
                <option value="pending" style={{ background: COLORS.chipBg, color: COLORS.ink, fontWeight: 700 }}>{t("hanzi_is_pending", meaningDisplay)}</option>
              </select>
            )}
            <button
              type="button"
              onClick={handleExportExcel}
              className="seal-btn"
              style={{ ...sealBtnStyle, padding: "8px 16px", fontSize: 13, flex: "none" }}
            >
              {t("hanzi_export_excel", meaningDisplay)}
            </button>
          </div>
          {exportMessage && (
            <div
              style={{
                textAlign: "center",
                fontSize: 12.5,
                fontWeight: 600,
                marginBottom: 10,
                color: exportMessage.type === "error" ? COLORS.error : COLORS.bamboo,
              }}
            >
              {exportMessage.text}
            </div>
          )}
          <div style={{ fontSize: 11.5, color: COLORS.inkSoft, textAlign: "center", marginBottom: 16 }}>
            {t("vocab_count", meaningDisplay, filtered.length, wordList.length)}
          </div>

          <div style={{ display: "grid", gridTemplateColumns: "repeat(auto-fill, minmax(170px, 1fr))", gap: 10 }}>
            {filtered.map((w, idx) => (
              <WordChip
                key={`${w.word}-${idx}`}
                w={w}
                characterList={characterList}
                findBushou={findBushou}
                allLists={allLists}
                onAddWord={onAddWord}
                onDeleteWord={onDeleteWord}
                onDeleteWordFromOfficial={onDeleteWordFromOfficial}
                isAdmin={isAdmin}
                isOfficial={officialWordKeys ? officialWordKeys.has(w.word) : false}
                hasOverride={overrideWordKeys ? overrideWordKeys.has(w.word) : false}
                onPromoteWord={onPromoteWord}
                onWithdrawWord={onWithdrawWord}
                meaningDisplay={meaningDisplay}
              />
            ))}
          </div>

          {filtered.length === 0 && (
            <div style={{ textAlign: "center", color: COLORS.inkSoft, fontSize: 13, padding: 20 }}>Không tìm thấy từ nào phù hợp.</div>
          )}
        </>
      )}
    </div>
  );
}

/* ---------- A single saved word: view mode (compact chip) + edit mode
   (expands into a small inline form). Saving re-upserts the same word via
   onAddWord, which already overwrites on conflict — same pattern as
   character and radical editing. ---------- */
function WordChip({ w, characterList, findBushou, allLists, onAddWord, onDeleteWord, onDeleteWordFromOfficial, isAdmin, isOfficial, hasOverride, onPromoteWord, onWithdrawWord, meaningDisplay }) {
  const [mode, setMode] = useState("view"); // view | edit
  const [zoomed, setZoomed] = useState(false);
  const [pinyin, setPinyin] = useState(w.pinyin);
  const [meaning, setMeaning] = useState(w.meaning);
  const [meaningVi, setMeaningVi] = useState(w.meaning_vi || "");
  const [sv, setSv] = useState(w.sv || "");
  const [selectedLists, setSelectedLists] = useState(w.lists || []);
  const [listTypeahead, setListTypeahead] = useState("");
  const [defaultStatus, setDefaultStatus] = useState("idle"); // idle | working | error

  async function handleToggleDefault() {
    setDefaultStatus("working");
    try {
      if (hasOverride || !isOfficial) await onPromoteWord(w);
      else await onWithdrawWord(w);
      setDefaultStatus("idle");
    } catch (e) {
      setDefaultStatus("error");
      setTimeout(() => setDefaultStatus("idle"), 2500);
    }
  }

  function startEdit() {
    setPinyin(w.pinyin);
    setMeaning(w.meaning);
    setMeaningVi(w.meaning_vi || "");
    setSv(w.sv || "");
    setSelectedLists(w.lists || []);
    setListTypeahead("");
    setMode("edit");
  }

  function addList(name) {
    const trimmed = name.trim();
    if (!trimmed) return;
    setSelectedLists((prev) => (prev.includes(trimmed) ? prev : [...prev, trimmed]));
    setListTypeahead("");
  }

  function toggleList(name) {
    setSelectedLists((prev) => (prev.includes(name) ? prev.filter((l) => l !== name) : [...prev, name]));
  }

  function removeList(name) {
    setSelectedLists((prev) => prev.filter((l) => l !== name));
  }

  function saveEdit() {
    if (!pinyin.trim() || !meaning.trim()) return;
    const lists = selectedLists
      .map((l) => l.trim())
      .filter(Boolean);
    onAddWord &&
      onAddWord({
        ...w,
        pinyin: pinyin.trim(),
        meaning: meaning.trim(),
        meaning_vi: meaningVi.trim(),
        sv: sv.trim(),
        lists: lists.length > 0 ? lists : ["Chưa phân loại"],
      });
    setMode("view");
  }

  if (mode === "edit") {
    return (
      <div
        style={{
          padding: "10px 12px",
          borderRadius: 6,
          border: `1.5px solid ${COLORS.gold}`,
          background: COLORS.card,
          fontSize: 12,
          width: "100%",
          boxSizing: "border-box",
          textAlign: "left",
        }}
      >
        <div style={{ fontFamily: "'Noto Serif SC', 'STKaiti', 'Kaiti SC', serif", fontSize: 18, marginBottom: 6 }}>{w.word}</div>
        <label style={{ fontSize: 10, color: COLORS.inkSoft, display: "block", marginBottom: 2 }}>Pinyin</label>
        <input value={pinyin} onChange={(e) => setPinyin(e.target.value)} style={{ ...inputStyle, width: "100%", marginBottom: 6, fontSize: 12, boxSizing: "border-box" }} />
        <label style={{ fontSize: 10, color: COLORS.inkSoft, display: "block", marginBottom: 2 }}>Nghĩa (English)</label>
        <input value={meaning} onChange={(e) => setMeaning(e.target.value)} style={{ ...inputStyle, width: "100%", marginBottom: 6, fontSize: 12, boxSizing: "border-box" }} />
        <label style={{ fontSize: 10, color: COLORS.inkSoft, display: "block", marginBottom: 2 }}>Nghĩa (Tiếng Việt)</label>
        <input value={meaningVi} onChange={(e) => setMeaningVi(e.target.value)} style={{ ...inputStyle, width: "100%", marginBottom: 6, fontSize: 12, boxSizing: "border-box" }} />
        <label style={{ fontSize: 10, color: COLORS.inkSoft, display: "block", marginBottom: 2 }}>Hán Việt</label>
        <input value={sv} onChange={(e) => setSv(e.target.value)} style={{ ...inputStyle, width: "100%", marginBottom: 6, fontSize: 12, boxSizing: "border-box" }} />
        <label style={{ fontSize: 10, color: COLORS.inkSoft, display: "block", marginBottom: 2 }}>Danh sách</label>
        {selectedLists.length > 0 && (
          <div style={{ display: "flex", flexWrap: "wrap", gap: 4, marginBottom: 6 }}>
            {selectedLists.map((l) => (
              <span
                key={l}
                style={{
                  display: "inline-flex",
                  alignItems: "center",
                  gap: 3,
                  fontSize: 11,
                  padding: "2px 5px 2px 8px",
                  borderRadius: 999,
                  border: `1px solid ${COLORS.seal}`,
                  background: "rgba(49,112,142,0.08)",
                  color: COLORS.seal,
                }}
              >
                {l}
                <button
                  type="button"
                  onClick={() => removeList(l)}
                  style={{ background: "none", border: "none", color: COLORS.seal, cursor: "pointer", fontSize: 11, lineHeight: 1, padding: 0 }}
                >
                  ✕
                </button>
              </span>
            ))}
          </div>
        )}
        <div style={{ display: "flex", gap: 4, marginBottom: 6 }}>
          <input
            value={listTypeahead}
            onChange={(e) => setListTypeahead(e.target.value)}
            onKeyDown={(e) => {
              if (e.key === "Enter") {
                e.preventDefault();
                addList(listTypeahead);
              }
            }}
            placeholder="+ danh sách"
            style={{ ...inputStyle, fontSize: 12, padding: "5px 8px" }}
          />
          <button
            type="button"
            onClick={() => addList(listTypeahead)}
            className="ghost-btn"
            style={{ ...ghostBtnStyle, padding: "5px 8px", fontSize: 11.5, flex: "none" }}
          >
            Thêm
          </button>
        </div>
        {allLists && allLists.length > 0 && (
          <div style={{ display: "flex", flexWrap: "wrap", gap: 4, marginBottom: 8 }}>
            {allLists.map((l) => (
              <button
                key={l}
                type="button"
                onClick={() => toggleList(l)}
                style={{
                  fontSize: 10.5,
                  padding: "2px 7px",
                  borderRadius: 999,
                  border: `1px solid ${selectedLists.includes(l) ? COLORS.seal : COLORS.grid}`,
                  background: selectedLists.includes(l) ? "rgba(49,112,142,0.08)" : "transparent",
                  color: selectedLists.includes(l) ? COLORS.seal : COLORS.inkSoft,
                  cursor: "pointer",
                }}
              >
                {selectedLists.includes(l) ? "✓ " : ""}
                {displayListName(l, meaningDisplay)}
              </button>
            ))}
          </div>
        )}
        <div style={{ display: "flex", gap: 6 }}>
          <button type="button" onClick={saveEdit} className="seal-btn" style={{ ...sealBtnStyle, padding: "5px 12px", fontSize: 11.5 }}>
            Lưu
          </button>
          <button type="button" onClick={() => setMode("view")} className="ghost-btn" style={{ ...ghostBtnStyle, padding: "5px 12px", fontSize: 11.5 }}>
            Hủy
          </button>
        </div>
      </div>
    );
  }

  return (
    <div
      style={{
        padding: "10px 9px",
        borderRadius: 6,
        border: `1px solid ${COLORS.hairline}`,
        background: COLORS.card,
        fontSize: 12,
        textAlign: "center",
        position: "relative",
      }}
    >
      <button
        type="button"
        onClick={() => setZoomed(true)}
        title="Phóng to để xem chi tiết"
        style={{
          position: "absolute",
          top: 6,
          left: 6,
          width: 18,
          height: 18,
          lineHeight: "16px",
          padding: 0,
          fontSize: 10,
          border: `1px solid ${COLORS.hairline}`,
          borderRadius: "50%",
          background: COLORS.chipBg,
          color: COLORS.sealDark,
          cursor: "pointer",
        }}
      >
        🔍
      </button>

      <div style={{ position: "absolute", top: 6, right: 6, display: "flex", gap: 4 }}>
        {(isAdmin || hasOverride || !isOfficial) && (
        <>
        <button
          type="button"
          onClick={startEdit}
          style={{ background: "none", border: "none", color: COLORS.gold, cursor: "pointer", fontSize: 12, padding: 0 }}
          title="Sửa từ này"
        >
          ✎
        </button>
        <button
          type="button"
          onClick={async () => {
            if (isAdmin && isOfficial) {
              if (!window.confirm(`Xóa "${w.word}" khỏi dữ liệu mặc định cho MỌI người dùng?`)) return;
              if (onDeleteWordFromOfficial) {
                try {
                  await onDeleteWordFromOfficial(w.word);
                } catch (e) {
                  console.error("Could not delete from default:", e);
                  alert(
                    `Không thể xóa "${w.word}" khỏi dữ liệu mặc định (lỗi: ${e && e.message ? e.message : "không rõ"}). Từ sẽ chỉ được ẩn khỏi tài khoản của bạn — vẫn hiển thị cho người dùng khác. Vui lòng kiểm tra lại trong Supabase.`
                  );
                }
              }
            }
            onDeleteWord && onDeleteWord(w.word);
          }}
          style={{ background: "none", border: "none", color: COLORS.error, cursor: "pointer", fontSize: 12, padding: 0 }}
          title="Xóa từ này"
        >
          ✕
        </button>
        </>
        )}
        {isOfficial && (
          <SuggestRevisionButton contentType="word" itemKey={w.word} meaningDisplay={meaningDisplay} compact />
        )}
      </div>

      <div
        onClick={() => setZoomed(true)}
        title="Bấm để phóng to"
        style={{ fontFamily: "'Noto Serif SC', 'STKaiti', 'Kaiti SC', serif", fontSize: 20, color: COLORS.ink, marginBottom: 2, marginTop: 6, cursor: "pointer" }}
      >
        {w.word}
      </div>
      <div style={{ color: COLORS.sealDark, fontSize: 11.5 }}>{w.pinyin}</div>
      <div style={{ marginTop: 4 }}>
        <MeaningBoxes meaning={w.meaning} meaningVi={w.meaning_vi} meaningDisplay={meaningDisplay} />
      </div>
      {w.sv && meaningDisplay !== "en" && (
        <div style={{ color: COLORS.bamboo, fontSize: 11.5, marginTop: 4, fontWeight: 600 }}>HV: {w.sv}</div>
      )}
      {isAdmin && (
        <button
          type="button"
          onClick={handleToggleDefault}
          disabled={defaultStatus === "working"}
          title={
            hasOverride || !isOfficial
              ? "Đặt/cập nhật làm dữ liệu mặc định cho mọi người dùng mới"
              : "Bấm để gỡ khỏi dữ liệu mặc định"
          }
          style={{
            marginTop: 6,
            fontSize: 10,
            padding: "3px 8px",
            borderRadius: 999,
            border: `1px solid ${!hasOverride && isOfficial ? COLORS.bamboo : COLORS.gold}`,
            background: !hasOverride && isOfficial ? "rgba(80,133,165,0.12)" : "rgba(80,133,165,0.06)",
            color: defaultStatus === "error" ? COLORS.error : !hasOverride && isOfficial ? COLORS.bamboo : COLORS.gold,
            cursor: defaultStatus === "working" ? "default" : "pointer",
            opacity: defaultStatus === "working" ? 0.6 : 1,
          }}
        >
          {defaultStatus === "working"
            ? "Đang xử lý…"
            : defaultStatus === "error"
            ? "✕ Lỗi, thử lại"
            : hasOverride && isOfficial
            ? "🔄 Cập nhật mặc định"
            : !hasOverride && isOfficial
            ? "★ Đang là mặc định"
            : "⭐ Đặt làm mặc định"}
        </button>
      )}

      {characterList && (
        <div style={{ marginTop: 8, paddingTop: 8, borderTop: `1px dashed ${COLORS.grid}`, textAlign: "left" }}>
          {Array.from(new Set(Array.from(w.word))).map((ch, i) => {
            const found = characterList.find((c) => c.char === ch);
            if (!found || !found.components || found.components.length === 0) return null;
            return (
              <div key={i} style={{ fontSize: 10.5, color: COLORS.inkSoft, marginBottom: 3 }}>
                <span style={{ fontFamily: "'Noto Serif SC', 'STKaiti', 'Kaiti SC', serif", color: COLORS.ink, fontSize: 13 }}>{ch}</span>
                {" = "}
                {found.components.map((comp, ci) => (
                  <span key={ci}>
                    <span
                      title={findBushou ? `${findBushou(comp).pinyin} · ${findBushou(comp).meaning} · HV: ${findBushou(comp).sv}` : undefined}
                      style={{ fontFamily: "'Noto Serif SC', 'STKaiti', 'Kaiti SC', serif", color: COLORS.gold, fontSize: 13 }}
                    >
                      {comp}
                    </span>
                    {ci < found.components.length - 1 ? " + " : ""}
                  </span>
                ))}
              </div>
            );
          })}
        </div>
      )}

      {zoomed && (
        <WordZoomModal w={w} characterList={characterList} findBushou={findBushou} onClose={() => setZoomed(false)} meaningDisplay={meaningDisplay} isOfficial={isOfficial} />
      )}
    </div>
  );
}

/* ---------- Full-screen study view for one word: one mizige box per
   character (same idea as the Play tab's multi-box build area), plus
   pinyin/meaning/Hán Việt and each character's own bushou breakdown. ---------- */
function WordZoomModal({ w, characterList, findBushou, onClose, meaningDisplay, isOfficial }) {
  const chars = Array.from(w.word);
  const boxSize = chars.length <= 2 ? 130 : chars.length === 3 ? 100 : 80;
  const [strokeChar, setStrokeChar] = useState(null);

  return (
    <div
      onClick={onClose}
      style={{
        position: "fixed",
        inset: 0,
        background: "rgba(30,28,10,0.55)",
        zIndex: 1000,
        display: "flex",
        alignItems: "center",
        justifyContent: "center",
        padding: 20,
      }}
    >
      <div
        onClick={(e) => e.stopPropagation()}
        style={{
          background: COLORS.card,
          borderRadius: 14,
          padding: "28px 24px",
          width: "90%",
          maxWidth: 480,
          maxHeight: "85vh",
          overflowY: "auto",
          textAlign: "center",
          position: "relative",
          boxShadow: "0 12px 40px rgba(0,0,0,0.35)",
        }}
      >
        <button
          type="button"
          onClick={onClose}
          title="Đóng"
          style={{
            position: "absolute",
            top: 12,
            right: 12,
            width: 28,
            height: 28,
            lineHeight: "26px",
            fontSize: 15,
            border: `1px solid ${COLORS.hairline}`,
            borderRadius: "50%",
            background: COLORS.chipBg,
            color: COLORS.inkSoft,
            cursor: "pointer",
          }}
        >
          ✕
        </button>

        <div style={{ fontSize: 11, color: COLORS.gold, fontWeight: 700, textTransform: "uppercase", letterSpacing: 0.6, marginBottom: 16 }}>
          {(w.lists || []).join(" · ")}
        </div>

        <div style={{ display: "flex", justifyContent: "center", gap: 10, marginBottom: 18, flexWrap: "wrap" }}>
          {chars.map((ch, i) => (
            <CharacterGrid key={i} state="revealed" size={boxSize}>
              <div style={{ fontFamily: "'Noto Serif SC', 'STKaiti', 'Kaiti SC', serif", fontSize: Math.round(boxSize * 0.6), color: COLORS.ink }}>
                {ch}
              </div>
            </CharacterGrid>
          ))}
        </div>

        <div style={{ marginBottom: 10 }}>
          <MeaningBoxes meaning={w.meaning} meaningVi={w.meaning_vi} meaningDisplay={meaningDisplay} large />
        </div>
        <div style={{ display: "flex", justifyContent: "center", gap: 22, marginBottom: 18, fontSize: 16, flexWrap: "wrap" }}>
          <span style={{ color: COLORS.sealDark }}>Pinyin: <strong>{w.pinyin}</strong></span>
          {w.sv && meaningDisplay !== "en" && <span style={{ color: COLORS.bamboo }}>Hán Việt: <strong>{w.sv}</strong></span>}
        </div>

        {characterList && (
          <div style={{ borderTop: `1px dashed ${COLORS.grid}`, paddingTop: 16 }}>
            <div style={{ fontSize: 11, color: COLORS.inkSoft, textTransform: "uppercase", letterSpacing: 0.5, fontWeight: 700, marginBottom: 12 }}>
              {t("components_breakdown_label", meaningDisplay)}
            </div>
            {Array.from(new Set(chars)).map((ch, i) => {
              const found = characterList.find((c) => c.char === ch);
              if (!found || !found.components || found.components.length === 0) return null;
              return (
                <div key={i} style={{ marginBottom: 14 }}>
                  <div style={{ fontSize: 13, color: COLORS.ink, marginBottom: 8, fontWeight: 600, display: "flex", alignItems: "center", justifyContent: "center", gap: 8 }}>
                    <span style={{ fontFamily: "'Noto Serif SC', 'STKaiti', 'Kaiti SC', serif", fontSize: 18 }}>{ch}</span>
                    <span>({found.pinyin}{meaningDisplay !== "vi" ? ` · ${found.meaning}` : ""})</span>
                    <button
                      type="button"
                      onClick={() => setStrokeChar(ch)}
                      title={t("view_stroke_order_of", meaningDisplay, ch)}
                      style={{
                        width: 20,
                        height: 20,
                        lineHeight: "18px",
                        padding: 0,
                        fontSize: 11,
                        border: `1px solid ${COLORS.hairline}`,
                        borderRadius: "50%",
                        background: COLORS.chipBg,
                        color: COLORS.sealDark,
                        cursor: "pointer",
                        flexShrink: 0,
                      }}
                    >
                      ✍️
                    </button>
                  </div>
                  <div style={{ display: "flex", justifyContent: "center", flexWrap: "wrap", gap: 8, marginBottom: 8 }}>
                    {found.components.map((comp, ci) => (
                      <Chip key={ci} info={findBushou(comp)} big disabled meaningDisplay={meaningDisplay} />
                    ))}
                  </div>
                  <div style={{ display: "flex", flexDirection: "column", gap: 3, textAlign: "left" }}>
                    {found.components.map((comp, ci) => {
                      const info = findBushou(comp);
                      return (
                        <div key={ci} style={{ fontSize: 12, color: COLORS.inkSoft }}>
                          <span style={{ fontFamily: "'Noto Serif SC', 'STKaiti', 'Kaiti SC', serif", fontSize: 15, color: COLORS.ink }}>{comp}</span>
                          {" — "}
                          {info.pinyin}
                          {meaningDisplay !== "vi" && ` · ${info.meaning}`}
                          {meaningDisplay !== "en" && ` · HV: ${info.sv}`}
                        </div>
                      );
                    })}
                  </div>
                </div>
              );
            })}
          </div>
        )}

        {strokeChar && <StrokeOrderModal char={strokeChar} onClose={() => setStrokeChar(null)} meaningDisplay={meaningDisplay} />}
        {isOfficial && (
          <div style={{ marginTop: 16, textAlign: "center" }}>
            <SuggestRevisionButton contentType="word" itemKey={w.word} meaningDisplay={meaningDisplay} />
          </div>
        )}
      </div>
    </div>
  );
}

/* ---------- List function: browse every character already in the database ---------- */
function CharacterListPanel({ characterList, bushouList, onDeleteCharacter, onDeleteCharacterFromOfficial, onUpdateCharacter, onAddBushou, isAdmin, officialCharKeys, overrideCharKeys, onPromoteCharacter, onWithdrawCharacter, checkListAccess, onViewPremium, meaningDisplay, userId }) {
  const [query, setQuery] = useState("");
  const [listFilter, setListFilter] = useState("Tất cả");
  const [defaultFilter, setDefaultFilter] = useState("all"); // all | official | pending
  const [exportMessage, setExportMessage] = useState(null);
  const [lockedListName, setLockedListName] = useState(null);

  const findBushou = (ch) =>
    bushouList.find((b) => b.char === ch) || { char: ch, pinyin: "—", meaning: "unknown", sv: "—" };

  const allLists = useMemo(() => {
    const set = new Set();
    characterList.forEach((c) => getLists(c).forEach((l) => set.add(l.trim())));
    return Array.from(set).sort((a, b) => a.localeCompare(b, "vi"));
  }, [characterList]);

  const sorted = useMemo(() => {
    return characterList
      .slice()
      .sort((a, b) => a.char.localeCompare(b.char, "zh"));
  }, [characterList]);

  const filtered = sorted.filter((c) => {
    if (listFilter !== "Tất cả" && !getLists(c).some((l) => l.trim() === listFilter)) return false;
    if (isAdmin && defaultFilter !== "all") {
      const isOfficial = officialCharKeys ? officialCharKeys.has(c.char) : false;
      const hasOverride = overrideCharKeys ? overrideCharKeys.has(c.char) : false;
      const isCleanlyPublished = isOfficial && !hasOverride;
      if (defaultFilter === "official" && !isCleanlyPublished) return false;
      if (defaultFilter === "pending" && isCleanlyPublished) return false;
    }
    const q = query.trim().toLowerCase();
    if (!q) return true;
    return (
      c.char.includes(q) ||
      c.pinyin.toLowerCase().includes(q) ||
      c.meaning.toLowerCase().includes(q) ||
      c.sv.toLowerCase().includes(q) ||
      (c.components || []).some((comp) => comp.includes(q))
    );
  });

  function handleExportExcel() {
    setExportMessage(null);
    try {
      if (filtered.length === 0) {
        setExportMessage({ type: "error", text: t("hanzi_export_none", meaningDisplay) });
        return;
      }
      const rows = filtered.map((c) => ({
        "Chữ Hán": c.char,
        Pinyin: c.pinyin,
        "Nghĩa (English)": c.meaning,
        "Nghĩa (Tiếng Việt)": c.meaning_vi || "",
        "Âm Hán Việt": c.sv,
        "Danh sách": getLists(c).join(", "),
        "Bộ thủ cấu thành": (c.components || []).join(" + "),
      }));
      const worksheet = XLSX.utils.json_to_sheet(rows);
      worksheet["!cols"] = [
        { wch: 10 },
        { wch: 14 },
        { wch: 32 },
        { wch: 32 },
        { wch: 16 },
        { wch: 20 },
        { wch: 24 },
      ];
      const workbook = XLSX.utils.book_new();
      XLSX.utils.book_append_sheet(workbook, worksheet, "Kho du lieu");
      const dateStr = new Date().toISOString().slice(0, 10);
      const suffix = listFilter !== "Tất cả" ? `-${listFilter}` : "";
      XLSX.writeFile(workbook, `kho-du-lieu-chu-han${suffix}-${dateStr}.xlsx`);
      setExportMessage({ type: "success", text: t("hanzi_export_success", meaningDisplay, filtered.length) });
    } catch (err) {
      console.error("Export to Excel failed:", err);
      setExportMessage({ type: "error", text: t("hanzi_export_fail", meaningDisplay) });
    }
  }

  return (
    <div>
      <div style={{ fontSize: 12.5, fontWeight: 600, color: COLORS.gold, marginBottom: 12, textTransform: "uppercase", letterSpacing: 0.8, textAlign: "center" }}>
        {userId ? t("hanzi_panel_title", meaningDisplay) : t("hanzi_panel_title_guest", meaningDisplay)}
      </div>

      <div style={{ display: "flex", justifyContent: "center", gap: 8, flexWrap: "wrap", marginBottom: 10 }}>
        <input
          value={query}
          onChange={(e) => setQuery(e.target.value)}
          placeholder={t("hanzi_search_placeholder", meaningDisplay)}
            style={{ ...inputStyle, width: 300, textAlign: "center" }}
          />
          <select
            value={listFilter}
            onChange={(e) => {
              const next = e.target.value;
              if (!isAdmin && next !== "Tất cả" && checkListAccess && !checkListAccess(next)) {
                setLockedListName(next);
                return;
              }
              setListFilter(next);
            }}
            style={{ ...selectStyle, width: 170, flex: "none" }}
          >
            <option value="Tất cả" style={{ background: COLORS.chipBg, color: COLORS.ink, fontWeight: 700 }}>{t("play_all_lists", meaningDisplay)}</option>
            {allLists.map((l) => (
              <option key={l} value={l} style={{ background: COLORS.chipBg, color: COLORS.ink, fontWeight: 700 }}>
                {!isAdmin && checkListAccess && !checkListAccess(l) ? `🔒 ${l}` : displayListName(l, meaningDisplay)}
              </option>
            ))}
          </select>
          {lockedListName && (
            <ListLockedModal listName={lockedListName} onClose={() => setLockedListName(null)} onViewPremium={onViewPremium} />
          )}
          {isAdmin && (
            <select value={defaultFilter} onChange={(e) => setDefaultFilter(e.target.value)} style={{ ...selectStyle, width: 190, flex: "none" }}>
              <option value="all" style={{ background: COLORS.chipBg, color: COLORS.ink, fontWeight: 700 }}>{t("hanzi_all_status", meaningDisplay)}</option>
              <option value="official" style={{ background: COLORS.chipBg, color: COLORS.ink, fontWeight: 700 }}>{t("hanzi_is_official", meaningDisplay)}</option>
              <option value="pending" style={{ background: COLORS.chipBg, color: COLORS.ink, fontWeight: 700 }}>{t("hanzi_is_pending", meaningDisplay)}</option>
            </select>
          )}
          <button
            type="button"
            onClick={handleExportExcel}
            className="seal-btn"
            style={{ ...sealBtnStyle, padding: "8px 16px", fontSize: 13, flex: "none" }}
          >
            {t("hanzi_export_excel", meaningDisplay)}
          </button>
        </div>
        {exportMessage && (
          <div
            style={{
              textAlign: "center",
              fontSize: 12.5,
              fontWeight: 600,
              marginBottom: 10,
              color: exportMessage.type === "error" ? COLORS.error : COLORS.bamboo,
            }}
          >
            {exportMessage.text}
          </div>
        )}
        <div style={{ fontSize: 12.5, color: COLORS.inkSoft, textAlign: "center", marginBottom: 16 }}>
          {t("hanzi_count", meaningDisplay, filtered.length, characterList.length)}
        </div>

        <div style={{ display: "grid", gridTemplateColumns: "repeat(auto-fill, minmax(170px, 1fr))", gap: 12 }}>
          {filtered.map((c) => (
            <CharacterCard
              key={c.char}
              c={c}
              bushouList={bushouList}
              findBushou={findBushou}
              onDeleteCharacter={onDeleteCharacter}
              onDeleteCharacterFromOfficial={onDeleteCharacterFromOfficial}
              onUpdateCharacter={onUpdateCharacter}
              onAddBushou={onAddBushou}
              allLists={allLists}
              isAdmin={isAdmin}
              isOfficial={officialCharKeys ? officialCharKeys.has(c.char) : false}
              hasOverride={overrideCharKeys ? overrideCharKeys.has(c.char) : false}
              onPromoteCharacter={onPromoteCharacter}
              onWithdrawCharacter={onWithdrawCharacter}
              meaningDisplay={meaningDisplay}
            />
          ))}
        </div>

        {filtered.length === 0 && (
          <div style={{ textAlign: "center", color: COLORS.inkSoft, fontSize: 13, padding: 30 }}>
            Không tìm thấy chữ nào phù hợp.
          </div>
        )}
    </div>
  );
}

/* ---------- A single character card: view mode, edit mode, delete confirm ---------- */
function CharacterCard({ c, bushouList, findBushou, onDeleteCharacter, onDeleteCharacterFromOfficial, onUpdateCharacter, onAddBushou, allLists, isAdmin, isOfficial, hasOverride, onPromoteCharacter, onWithdrawCharacter, meaningDisplay }) {
  const [mode, setMode] = useState("view"); // view | edit | confirmDelete
  const [zoomed, setZoomed] = useState(false);
  const [meaning, setMeaning] = useState(c.meaning);
  const [meaningVi, setMeaningVi] = useState(c.meaning_vi || "");
  const [pinyin, setPinyin] = useState(c.pinyin);
  const [sv, setSv] = useState(c.sv);
  const [lists, setLists] = useState(getLists(c));
  const [listTypeahead, setListTypeahead] = useState("");
  const [components, setComponents] = useState(c.components || []);
  const [compInput, setCompInput] = useState("");
  const [newCompDraft, setNewCompDraft] = useState(null);
  const [ncPinyin, setNcPinyin] = useState("");
  const [ncMeaning, setNcMeaning] = useState("");
  const [ncSv, setNcSv] = useState("");
  const [defaultStatus, setDefaultStatus] = useState("idle"); // idle | working | error

  async function handleToggleDefault() {
    setDefaultStatus("working");
    try {
      if (hasOverride || !isOfficial) await onPromoteCharacter(c);
      else await onWithdrawCharacter(c);
      setDefaultStatus("idle");
    } catch (e) {
      setDefaultStatus("error");
      setTimeout(() => setDefaultStatus("idle"), 2500);
    }
  }

  function startEdit() {
    setMeaning(c.meaning);
    setMeaningVi(c.meaning_vi || "");
    setPinyin(c.pinyin);
    setSv(c.sv);
    setLists(getLists(c));
    setListTypeahead("");
    setComponents(c.components || []);
    setCompInput("");
    setNewCompDraft(null);
    setMode("edit");
  }

  function cancelEdit() {
    setMode("view");
  }

  function addListToEdit(name) {
    const trimmed = name.trim();
    if (!trimmed) return;
    setLists((prev) => (prev.includes(trimmed) ? prev : [...prev, trimmed]));
    setListTypeahead("");
  }

  function removeListFromEdit(name) {
    setLists((prev) => prev.filter((l) => l !== name));
  }

  function saveEdit() {
    if (!meaning.trim() || !pinyin.trim() || !sv.trim()) return;
    onUpdateCharacter &&
      onUpdateCharacter(c.char, {
        meaning: meaning.trim(),
        meaning_vi: meaningVi.trim(),
        pinyin: pinyin.trim(),
        sv: sv.trim(),
        lists: lists.length > 0 ? lists : ["Chưa phân loại"],
        list: undefined,
        components,
      });
    setMode("view");
  }

  function addComponentManually() {
    const ch = compInput.trim();
    if (!ch) return;
    const known = bushouList.find((b) => b.char === ch);
    if (known) {
      setComponents((prev) => [...prev, ch]);
      setCompInput("");
    } else {
      setNewCompDraft({ char: ch });
      setNcPinyin("");
      setNcMeaning("");
      setNcSv("");
    }
  }

  function confirmNewComponent() {
    if (!newCompDraft || !ncPinyin.trim() || !ncMeaning.trim() || !ncSv.trim()) return;
    if (onAddBushou) {
      onAddBushou({ char: newCompDraft.char, pinyin: ncPinyin.trim(), meaning: ncMeaning.trim(), sv: ncSv.trim() });
    }
    setComponents((prev) => [...prev, newCompDraft.char]);
    setNewCompDraft(null);
    setCompInput("");
  }

  function removeComponent(idx) {
    setComponents((prev) => prev.filter((_, i) => i !== idx));
  }

  return (
    <div
      style={{
        background: COLORS.card,
        border: `1px solid ${COLORS.hairline}`,
        borderRadius: 11,
        padding: "12px 10px",
        textAlign: "center",
        position: "relative",
      }}
    >
      {mode !== "edit" && (
        <button
          type="button"
          onClick={() => setZoomed(true)}
          title="Phóng to để xem chi tiết"
          style={{
            position: "absolute",
            top: 6,
            left: 6,
            width: 20,
            height: 20,
            lineHeight: "18px",
            padding: 0,
            fontSize: 11,
            border: `1px solid ${COLORS.hairline}`,
            borderRadius: "50%",
            background: COLORS.chipBg,
            color: COLORS.sealDark,
            cursor: "pointer",
          }}
        >
          🔍
        </button>
      )}

      {mode !== "edit" && (
        <div style={{ position: "absolute", top: 6, right: 6, display: "flex", gap: 4 }}>
          {(isAdmin || hasOverride || !isOfficial) && (
            <>
          <button
            type="button"
            onClick={startEdit}
            title="Sửa chữ này"
            style={{
              width: 20,
              height: 20,
              lineHeight: "18px",
              padding: 0,
              fontSize: 11,
              border: `1px solid ${COLORS.hairline}`,
              borderRadius: "50%",
              background: COLORS.chipBg,
              color: COLORS.gold,
              cursor: "pointer",
            }}
          >
            ✎
          </button>
          <button
            type="button"
            onClick={() => setMode(mode === "confirmDelete" ? "view" : "confirmDelete")}
            title="Xóa chữ này"
            style={{
              width: 20,
              height: 20,
              lineHeight: "18px",
              padding: 0,
              fontSize: 12,
              border: `1px solid ${COLORS.hairline}`,
              borderRadius: "50%",
              background: COLORS.chipBg,
              color: COLORS.error,
              cursor: "pointer",
            }}
          >
            ✕
          </button>
            </>
          )}
          {isOfficial && (
            <SuggestRevisionButton contentType="char" itemKey={c.char} meaningDisplay={meaningDisplay} compact />
          )}
        </div>
      )}

      {mode === "edit" ? (
        <div style={{ textAlign: "left" }}>
          <div style={{ fontFamily: "'Noto Serif SC', 'STKaiti', 'Kaiti SC', serif", fontSize: 26, color: COLORS.ink, textAlign: "center", marginBottom: 8 }}>
            {c.char}
          </div>
          <label style={{ fontSize: 10, color: COLORS.inkSoft, display: "block", marginBottom: 2 }}>Nghĩa (English)</label>
          <input value={meaning} onChange={(e) => setMeaning(e.target.value)} style={{ ...inputStyle, width: "100%", marginBottom: 6, fontSize: 12.5 }} />
          <label style={{ fontSize: 10, color: COLORS.inkSoft, display: "block", marginBottom: 2 }}>Nghĩa (Tiếng Việt)</label>
          <input value={meaningVi} onChange={(e) => setMeaningVi(e.target.value)} style={{ ...inputStyle, width: "100%", marginBottom: 6, fontSize: 12.5 }} />
          <label style={{ fontSize: 10, color: COLORS.inkSoft, display: "block", marginBottom: 2 }}>Pinyin</label>
          <input value={pinyin} onChange={(e) => setPinyin(e.target.value)} style={{ ...inputStyle, width: "100%", marginBottom: 6, fontSize: 12.5 }} />
          <label style={{ fontSize: 10, color: COLORS.inkSoft, display: "block", marginBottom: 2 }}>Hán Việt</label>
          <input value={sv} onChange={(e) => setSv(e.target.value)} style={{ ...inputStyle, width: "100%", marginBottom: 6, fontSize: 12.5 }} />
          <label style={{ fontSize: 10, color: COLORS.inkSoft, display: "block", marginBottom: 2 }}>Danh sách (có thể nhiều)</label>
          {lists.length > 0 && (
            <div style={{ display: "flex", flexWrap: "wrap", gap: 4, marginBottom: 4 }}>
              {lists.map((l) => (
                <span
                  key={l}
                  style={{
                    display: "inline-flex",
                    alignItems: "center",
                    gap: 3,
                    fontSize: 11,
                    padding: "2px 5px 2px 8px",
                    borderRadius: 999,
                    border: `1px solid ${COLORS.seal}`,
                    background: "rgba(168,72,47,0.08)",
                    color: COLORS.seal,
                  }}
                >
                  {l}
                  <button
                    type="button"
                    onClick={() => removeListFromEdit(l)}
                    style={{ background: "none", border: "none", color: COLORS.seal, cursor: "pointer", fontSize: 11, lineHeight: 1, padding: 0 }}
                  >
                    ✕
                  </button>
                </span>
              ))}
            </div>
          )}
          <div style={{ display: "flex", gap: 4, marginBottom: 8 }}>
            <input
              value={listTypeahead}
              onChange={(e) => setListTypeahead(e.target.value)}
              onKeyDown={(e) => {
                if (e.key === "Enter") {
                  e.preventDefault();
                  addListToEdit(listTypeahead);
                }
              }}
              placeholder="+ danh sách"
              style={{ ...inputStyle, fontSize: 12.5, padding: "5px 8px" }}
            />
            <button
              type="button"
              onClick={() => addListToEdit(listTypeahead)}
              className="ghost-btn"
              style={{ ...ghostBtnStyle, padding: "5px 8px", fontSize: 11.5, flex: "none" }}
            >
              Thêm
            </button>
          </div>

          {allLists && allLists.filter((l) => !lists.includes(l)).length > 0 && (
            <select
              value=""
              onChange={(e) => {
                if (e.target.value) addListToEdit(e.target.value);
              }}
              style={{ ...selectStyle, width: "100%", fontSize: 12.5, padding: "5px 8px", marginBottom: 8 }}
            >
              <option value="" style={{ background: COLORS.chipBg, color: COLORS.ink, fontWeight: 700 }}>— Recent Lists —</option>
              {allLists
                .filter((l) => !lists.includes(l))
                .map((l) => (
                  <option key={l} value={l} style={{ background: COLORS.chipBg, color: COLORS.ink, fontWeight: 700 }}>
                    {l}
                  </option>
                ))}
            </select>
          )}

          <label style={{ fontSize: 10, color: COLORS.inkSoft, display: "block", marginBottom: 4 }}>{t("components_breakdown_label", meaningDisplay)}</label>
          {components.length > 0 && (
            <div style={{ display: "flex", flexWrap: "wrap", gap: 4, marginBottom: 6 }}>
              {components.map((ch, i) => (
                <div key={i} style={{ display: "flex", alignItems: "center", gap: 2 }}>
                  <span
                    style={{
                      fontFamily: "'Noto Serif SC', 'STKaiti', 'Kaiti SC', serif",
                      fontSize: 14,
                      color: COLORS.gold,
                      padding: "1px 5px",
                      border: `1px solid ${COLORS.hairline}`,
                      borderRadius: 4,
                      background: COLORS.chipBg,
                    }}
                  >
                    {ch}
                  </span>
                  <button type="button" onClick={() => removeComponent(i)} style={{ ...smallXStyle, fontSize: 11 }}>
                    ✕
                  </button>
                </div>
              ))}
            </div>
          )}
          <div style={{ display: "flex", gap: 4, marginBottom: 6 }}>
            <input
              value={compInput}
              onChange={(e) => setCompInput(e.target.value)}
              placeholder="+ bộ thủ"
              style={{ ...inputStyle, fontFamily: "'Noto Serif SC', 'STKaiti', 'Kaiti SC', serif", fontSize: 12.5, padding: "5px 8px" }}
              onKeyDown={(e) => {
                if (e.key === "Enter") {
                  e.preventDefault();
                  addComponentManually();
                }
              }}
            />
            <button
              type="button"
              onClick={addComponentManually}
              className="ghost-btn"
              style={{ ...ghostBtnStyle, padding: "5px 8px", fontSize: 11.5, flex: "none" }}
            >
              Thêm
            </button>
          </div>
          {newCompDraft && (
            <div style={{ padding: 8, background: "rgba(80,133,165,0.08)", borderRadius: 6, border: `1px dashed ${COLORS.gold}`, marginBottom: 6 }}>
              <div style={{ fontSize: 11, marginBottom: 6 }}>
                Bộ thủ <strong style={{ fontFamily: "'Noto Serif SC', 'STKaiti', 'Kaiti SC', serif" }}>{newCompDraft.char}</strong> chưa có — điền thông tin:
              </div>
              <div style={{ display: "flex", gap: 4, flexWrap: "wrap" }}>
                <input value={ncPinyin} onChange={(e) => setNcPinyin(e.target.value)} placeholder="pinyin" style={{ ...inputStyle, fontSize: 11.5, padding: "4px 6px", width: 60 }} />
                <input value={ncMeaning} onChange={(e) => setNcMeaning(e.target.value)} placeholder="meaning" style={{ ...inputStyle, fontSize: 11.5, padding: "4px 6px", width: 80 }} />
                <input value={ncSv} onChange={(e) => setNcSv(e.target.value)} placeholder="Hán Việt" style={{ ...inputStyle, fontSize: 11.5, padding: "4px 6px", width: 70 }} />
                <button type="button" onClick={confirmNewComponent} className="seal-btn" style={{ ...sealBtnStyle, padding: "4px 8px", fontSize: 11 }}>
                  OK
                </button>
              </div>
            </div>
          )}

          <div style={{ display: "flex", justifyContent: "center", gap: 6, marginTop: 8 }}>
            <button type="button" onClick={saveEdit} className="seal-btn" style={{ ...sealBtnStyle, padding: "6px 14px", fontSize: 12 }}>
              Lưu
            </button>
            <button type="button" onClick={cancelEdit} className="ghost-btn" style={{ ...ghostBtnStyle, padding: "6px 14px", fontSize: 12 }}>
              Hủy
            </button>
          </div>
        </div>
      ) : (
        <>
          <div
            onClick={() => setZoomed(true)}
            title="Bấm để phóng to"
            style={{ fontFamily: "'Noto Serif SC', 'STKaiti', 'Kaiti SC', serif", fontSize: 32, color: COLORS.ink, cursor: "pointer", marginTop: 10 }}
          >
            {c.char}
          </div>
          <div style={{ fontSize: 12.5, color: COLORS.sealDark, marginTop: 4 }}>{c.pinyin}</div>
          <div style={{ marginTop: 4 }}>
            <MeaningBoxes meaning={c.meaning} meaningVi={c.meaning_vi} meaningDisplay={meaningDisplay} />
          </div>
          {meaningDisplay !== "en" && (
            <div style={{ fontSize: 11.5, color: COLORS.bamboo, marginTop: 4, fontWeight: 600 }}>HV: {c.sv}</div>
          )}
          {isAdmin && (
            <button
              type="button"
              onClick={handleToggleDefault}
              disabled={defaultStatus === "working"}
              title={
                hasOverride || !isOfficial
                  ? "Đặt/cập nhật làm dữ liệu mặc định cho mọi người dùng mới"
                  : "Bấm để gỡ khỏi dữ liệu mặc định"
              }
              style={{
                marginTop: 6,
                fontSize: 10,
                padding: "3px 8px",
                borderRadius: 999,
                border: `1px solid ${!hasOverride && isOfficial ? COLORS.bamboo : COLORS.gold}`,
                background: !hasOverride && isOfficial ? "rgba(80,133,165,0.12)" : "rgba(80,133,165,0.06)",
                color: defaultStatus === "error" ? COLORS.error : !hasOverride && isOfficial ? COLORS.bamboo : COLORS.gold,
                cursor: defaultStatus === "working" ? "default" : "pointer",
                opacity: defaultStatus === "working" ? 0.6 : 1,
              }}
            >
              {defaultStatus === "working"
                ? "Đang xử lý…"
                : defaultStatus === "error"
                ? "✕ Lỗi, thử lại"
                : hasOverride && isOfficial
                ? "🔄 Cập nhật mặc định"
                : !hasOverride && isOfficial
                ? "★ Đang là mặc định"
                : "⭐ Đặt làm mặc định"}
            </button>
          )}
          {c.components && c.components.length > 0 && (
            <div
              style={{
                marginTop: 8,
                paddingTop: 8,
                borderTop: `1px dashed ${COLORS.grid}`,
                display: "flex",
                justifyContent: "center",
                flexWrap: "wrap",
                gap: 4,
              }}
            >
              {c.components.map((comp, i) => (
                <span
                  key={i}
                  title={`${findBushou(comp).pinyin} · ${findBushou(comp).meaning} · HV: ${findBushou(comp).sv}`}
                  style={{
                    fontFamily: "'Noto Serif SC', 'STKaiti', 'Kaiti SC', serif",
                    fontSize: 16,
                    color: COLORS.gold,
                    padding: "2px 5px",
                    border: `1px solid ${COLORS.hairline}`,
                    borderRadius: 4,
                    background: COLORS.chipBg,
                  }}
                >
                  {comp}
                </span>
              ))}
            </div>
          )}
          {mode === "confirmDelete" && (
            <div
              style={{
                marginTop: 8,
                paddingTop: 8,
                borderTop: `1px dashed ${COLORS.error}`,
                display: "flex",
                flexDirection: "column",
                gap: 6,
              }}
            >
              <div style={{ fontSize: 11, color: COLORS.error, fontWeight: 600 }}>
                {isAdmin && isOfficial
                  ? `Xóa "${c.char}" khỏi dữ liệu mặc định cho MỌI người dùng?`
                  : `Xóa chữ "${c.char}"?`}
              </div>
              <div style={{ display: "flex", justifyContent: "center", gap: 6 }}>
                <button
                  type="button"
                  onClick={async () => {
                    if (isAdmin && isOfficial && onDeleteCharacterFromOfficial) {
                      try {
                        await onDeleteCharacterFromOfficial(c.char);
                      } catch (e) {
                        console.error("Could not delete from default:", e);
                        alert(
                          `Không thể xóa "${c.char}" khỏi dữ liệu mặc định (lỗi: ${e && e.message ? e.message : "không rõ"}). Chữ sẽ chỉ được ẩn khỏi tài khoản của bạn — vẫn hiển thị cho người dùng khác. Vui lòng kiểm tra lại trong Supabase.`
                        );
                      }
                    }
                    onDeleteCharacter && onDeleteCharacter(c.char);
                    setMode("view");
                  }}
                  style={{
                    background: COLORS.error,
                    border: "none",
                    color: "#FBF9EF",
                    borderRadius: 5,
                    padding: "4px 10px",
                    fontSize: 11.5,
                    cursor: "pointer",
                  }}
                >
                  Xóa
                </button>
                <button
                  type="button"
                  onClick={() => setMode("view")}
                  className="ghost-btn"
                  style={{ ...ghostBtnStyle, padding: "4px 10px", fontSize: 11.5 }}
                >
                  Hủy
                </button>
              </div>
            </div>
          )}
        </>
      )}

      {zoomed && <CharacterZoomModal c={c} findBushou={findBushou} onClose={() => setZoomed(false)} meaningDisplay={meaningDisplay} isOfficial={isOfficial} />}
    </div>
  );
}

/* ---------- Full-screen study view for one character: big glyph in the
   same mizige grid used during Play, plus pinyin/meaning/Hán Việt and each
   component's own details, all at a much larger size than the card. ---------- */
function CharacterZoomModal({ c, findBushou, onClose, meaningDisplay, isOfficial }) {
  const [strokeOrderOpen, setStrokeOrderOpen] = useState(false);
  return (
    <div
      onClick={onClose}
      style={{
        position: "fixed",
        inset: 0,
        background: "rgba(30,28,10,0.55)",
        zIndex: 1000,
        display: "flex",
        alignItems: "center",
        justifyContent: "center",
        padding: 20,
      }}
    >
      <div
        onClick={(e) => e.stopPropagation()}
        style={{
          background: COLORS.card,
          borderRadius: 14,
          padding: "28px 24px",
          width: "90%",
          maxWidth: 420,
          maxHeight: "85vh",
          overflowY: "auto",
          textAlign: "center",
          position: "relative",
          boxShadow: "0 12px 40px rgba(0,0,0,0.35)",
        }}
      >
        <button
          type="button"
          onClick={onClose}
          title="Đóng"
          style={{
            position: "absolute",
            top: 12,
            right: 12,
            width: 28,
            height: 28,
            lineHeight: "26px",
            fontSize: 15,
            border: `1px solid ${COLORS.hairline}`,
            borderRadius: "50%",
            background: COLORS.chipBg,
            color: COLORS.inkSoft,
            cursor: "pointer",
          }}
        >
          ✕
        </button>

        <div style={{ fontSize: 11, color: COLORS.gold, fontWeight: 700, textTransform: "uppercase", letterSpacing: 0.6, marginBottom: 16 }}>
          {getLists(c).join(" · ")}
        </div>

        <div style={{ display: "flex", justifyContent: "center", marginBottom: 18 }}>
          <CharacterGrid state="revealed" size={220}>
            <div style={{ fontFamily: "'Noto Serif SC', 'STKaiti', 'Kaiti SC', serif", fontSize: 130, color: COLORS.ink }}>{c.char}</div>
          </CharacterGrid>
        </div>

        <div style={{ marginBottom: 10 }}>
          <MeaningBoxes meaning={c.meaning} meaningVi={c.meaning_vi} meaningDisplay={meaningDisplay} large />
        </div>
        <div style={{ display: "flex", justifyContent: "center", gap: 22, marginBottom: 12, fontSize: 16, flexWrap: "wrap" }}>
          <span style={{ color: COLORS.sealDark }}>Pinyin: <strong>{c.pinyin}</strong></span>
          {meaningDisplay !== "en" && <span style={{ color: COLORS.bamboo }}>Hán Việt: <strong>{c.sv}</strong></span>}
        </div>

        <div style={{ marginBottom: 18 }}>
          <button
            type="button"
            onClick={() => setStrokeOrderOpen(true)}
            className="ghost-btn"
            style={{ ...ghostBtnStyle, borderColor: COLORS.sealDark, color: COLORS.sealDark, fontSize: 12.5 }}
          >
            {"✍️ "}{t("radicals_view_stroke_order", meaningDisplay)}
          </button>
        </div>

        {strokeOrderOpen && <StrokeOrderModal char={c.char} onClose={() => setStrokeOrderOpen(false)} meaningDisplay={meaningDisplay} />}

        {c.components && c.components.length > 0 && (
          <div style={{ borderTop: `1px dashed ${COLORS.grid}`, paddingTop: 16 }}>
            <div style={{ fontSize: 11, color: COLORS.inkSoft, textTransform: "uppercase", letterSpacing: 0.5, fontWeight: 700, marginBottom: 10 }}>
              {t("components_breakdown_label", meaningDisplay)}
            </div>
            <div style={{ display: "flex", justifyContent: "center", flexWrap: "wrap", gap: 10, marginBottom: 12 }}>
              {c.components.map((comp, i) => (
                <Chip key={i} info={findBushou(comp)} big disabled meaningDisplay={meaningDisplay} />
              ))}
            </div>
            <div style={{ display: "flex", flexDirection: "column", gap: 4, textAlign: "left" }}>
              {c.components.map((comp, i) => {
                const info = findBushou(comp);
                return (
                  <div key={i} style={{ fontSize: 12.5, color: COLORS.inkSoft }}>
                    <span style={{ fontFamily: "'Noto Serif SC', 'STKaiti', 'Kaiti SC', serif", fontSize: 16, color: COLORS.ink }}>{comp}</span>
                    {" — "}
                    {info.pinyin}
                    {meaningDisplay !== "vi" && ` · ${info.meaning}`}
                    {meaningDisplay !== "en" && ` · HV: ${info.sv}`}
                  </div>
                );
              })}
            </div>
          </div>
        )}
        {isOfficial && (
          <div style={{ marginTop: 16, textAlign: "center" }}>
            <SuggestRevisionButton contentType="char" itemKey={c.char} meaningDisplay={meaningDisplay} />
          </div>
        )}
      </div>
    </div>
  );
}

// A small, collapsible "flag an issue" control for shared/official cards.
// Expands into a short message box in place, rather than opening another
// modal on top of one that may already be open (e.g. inside a zoom modal).
function SuggestRevisionButton({ contentType, itemKey, meaningDisplay, compact }) {
  const [open, setOpen] = useState(false);
  const [message, setMessage] = useState("");
  const [guestEmail, setGuestEmail] = useState("");
  const [isGuest, setIsGuest] = useState(null); // null = still checking
  const [status, setStatus] = useState("idle"); // idle | sending | done | error

  async function handleOpen(e) {
    if (e) e.stopPropagation();
    setOpen(true);
    const {
      data: { user },
    } = await supabase.auth.getUser();
    setIsGuest(!user);
  }

  async function handleSubmit(e) {
    e.preventDefault();
    if (!message.trim()) return;
    setStatus("sending");
    try {
      // Signed-in users' email comes straight from their account. Guests
      // get an optional field, since there's no account email to pull.
      const {
        data: { user },
      } = await supabase.auth.getUser();
      const { error } = await supabase.from("card_suggestions").insert({
        content_type: contentType,
        item_key: itemKey,
        message: message.trim(),
        email: user ? user.email : guestEmail.trim() || null,
        user_id: user ? user.id : null,
      });
      if (error) throw error;
      setStatus("done");
      setMessage("");
      setGuestEmail("");
    } catch (err) {
      console.error("Could not submit suggestion:", err);
      setStatus("error");
    }
  }

  if (!open) {
    if (compact) {
      return (
        <button
          type="button"
          onClick={handleOpen}
          title={t("suggest_revision_button", meaningDisplay)}
          style={{
            width: 20,
            height: 20,
            lineHeight: "18px",
            padding: 0,
            fontSize: 11,
            border: `1px solid ${COLORS.hairline}`,
            borderRadius: "50%",
            background: COLORS.chipBg,
            color: COLORS.error,
            cursor: "pointer",
          }}
        >
          🚩
        </button>
      );
    }
    return (
      <button
        type="button"
        onClick={handleOpen}
        style={{ background: "none", border: "none", color: COLORS.metadata, fontSize: 11, cursor: "pointer", padding: 0, textDecoration: "underline" }}
      >
        🚩 {t("suggest_revision_button", meaningDisplay)}
      </button>
    );
  }

  if (status === "done") {
    return (
      <div
        style={
          compact
            ? { position: "absolute", top: 28, right: 6, zIndex: 5, background: COLORS.card, border: `1px solid ${COLORS.hairline}`, borderRadius: 8, padding: "8px 10px", width: 180, boxShadow: "0 2px 8px rgba(0,0,0,0.12)" }
            : {}
        }
      >
        <div style={{ fontSize: 11.5, color: COLORS.seal, fontWeight: 600 }}>{t("suggest_revision_thanks", meaningDisplay)}</div>
      </div>
    );
  }

  return (
    <form
      onSubmit={handleSubmit}
      onClick={(e) => compact && e.stopPropagation()}
      style={
        compact
          ? { position: "absolute", top: 28, right: 6, zIndex: 5, background: COLORS.card, border: `1px solid ${COLORS.hairline}`, borderRadius: 8, padding: "10px", width: 190, boxShadow: "0 2px 8px rgba(0,0,0,0.12)", textAlign: "left" }
          : { marginTop: 4 }
      }
    >
      <textarea
        value={message}
        onChange={(e) => setMessage(e.target.value)}
        placeholder={t("suggest_revision_placeholder", meaningDisplay)}
        rows={2}
        style={{ ...inputStyle, width: "100%", boxSizing: "border-box", fontSize: 12, resize: "vertical", marginBottom: isGuest ? 4 : 6 }}
      />
      {isGuest && (
        <input
          value={guestEmail}
          onChange={(e) => setGuestEmail(e.target.value)}
          placeholder={t("feedback_email_placeholder", meaningDisplay)}
          style={{ ...inputStyle, width: "100%", boxSizing: "border-box", fontSize: 12, marginBottom: 6 }}
        />
      )}
      <div style={{ display: "flex", gap: 6 }}>
        <button type="submit" className="seal-btn" style={{ ...sealBtnStyle, padding: "4px 12px", fontSize: 11.5 }} disabled={status === "sending"}>
          {status === "sending" ? t("feedback_sending", meaningDisplay) : t("suggest_revision_submit", meaningDisplay)}
        </button>
        <button type="button" onClick={() => setOpen(false)} className="ghost-btn" style={{ ...ghostBtnStyle, padding: "4px 12px", fontSize: 11.5 }}>
          {t("admin_deck_cancel", meaningDisplay)}
        </button>
      </div>
      {status === "error" && <div style={{ fontSize: 11, color: COLORS.error, marginTop: 4 }}>{t("feedback_error", meaningDisplay)}</div>}
    </form>
  );
}

function stripHtml(html) {
  const div = document.createElement("div");
  div.innerHTML = html;
  return (div.textContent || div.innerText || "").replace(/\s+/g, " ").trim();
}

// Wraps runs of Chinese characters in a span (class "cjk-enhanced") so they
// can get their own font and size, independent of surrounding Latin text --
// even when both sit in the same sentence. Walks actual DOM text nodes
// rather than regex-matching the raw HTML string, so it can't accidentally
// alter tag names or attribute values.
const CJK_RANGE = /[\u3400-\u9FFF\uF900-\uFAFF]+/g;
function enhanceCjkText(text) {
  if (!CJK_RANGE.test(text)) return null;
  CJK_RANGE.lastIndex = 0;
  const frag = document.createDocumentFragment();
  let lastIndex = 0;
  let match;
  while ((match = CJK_RANGE.exec(text))) {
    if (match.index > lastIndex) frag.appendChild(document.createTextNode(text.slice(lastIndex, match.index)));
    const span = document.createElement("span");
    span.className = "cjk-enhanced";
    span.textContent = match[0];
    frag.appendChild(span);
    lastIndex = match.index + match[0].length;
  }
  if (lastIndex < text.length) frag.appendChild(document.createTextNode(text.slice(lastIndex)));
  return frag;
}
function enhanceCjkInHtml(html) {
  const container = document.createElement("div");
  container.innerHTML = html;
  function walk(node) {
    if (node.nodeType === 3) {
      const frag = enhanceCjkText(node.nodeValue);
      if (frag) node.parentNode.replaceChild(frag, node);
    } else if (node.nodeType === 1) {
      Array.from(node.childNodes).forEach(walk);
    }
  }
  Array.from(container.childNodes).forEach(walk);
  return container.innerHTML;
}

function RichTextEditor({ value, onChange, placeholder }) {
  const editorRef = useRef(null);
  const lastValueRef = useRef(null); // sentinel -- guarantees the very first sync actually runs

  useEffect(() => {
    // Only push external value changes into the DOM (e.g. loading a
    // different post to edit) -- never on every keystroke, or the cursor
    // would jump to the start on each render.
    if (editorRef.current && value !== lastValueRef.current) {
      editorRef.current.innerHTML = value || "";
      lastValueRef.current = value;
    }
  }, [value]);

  function handleInput() {
    const html = editorRef.current.innerHTML;
    lastValueRef.current = html;
    onChange(html);
  }

  function exec(command, arg) {
    editorRef.current.focus();
    document.execCommand(command, false, arg);
    handleInput();
  }

  function insertLink() {
    const url = window.prompt("Link URL:");
    if (url) exec("createLink", url);
  }

  const toolbarBtnStyle = {
    background: "none",
    border: `1px solid ${COLORS.hairline}`,
    borderRadius: 6,
    padding: "5px 9px",
    fontSize: 12.5,
    color: COLORS.ink,
    cursor: "pointer",
  };

  return (
    <div>
      <div style={{ display: "flex", gap: 5, flexWrap: "wrap", marginBottom: 6 }}>
        <button type="button" onClick={() => exec("bold")} style={{ ...toolbarBtnStyle, fontWeight: 700 }}>B</button>
        <button type="button" onClick={() => exec("italic")} style={{ ...toolbarBtnStyle, fontStyle: "italic" }}>I</button>
        <button type="button" onClick={() => exec("underline")} style={{ ...toolbarBtnStyle, textDecoration: "underline" }}>U</button>
        <select
          onChange={(e) => {
            if (e.target.value) exec("fontSize", e.target.value);
            e.target.value = "";
          }}
          defaultValue=""
          style={{ ...toolbarBtnStyle, cursor: "pointer" }}
        >
          <option value="">Size</option>
          <option value="2">Small</option>
          <option value="3">Normal</option>
          <option value="5">Large</option>
          <option value="7">X-Large</option>
        </select>
        <button type="button" onClick={() => exec("justifyLeft")} style={toolbarBtnStyle}>⯇</button>
        <button type="button" onClick={() => exec("justifyCenter")} style={toolbarBtnStyle}>☰</button>
        <button type="button" onClick={() => exec("justifyRight")} style={toolbarBtnStyle}>⯈</button>
        <button type="button" onClick={() => exec("justifyFull")} style={toolbarBtnStyle}>≡</button>
        <button type="button" onClick={insertLink} style={toolbarBtnStyle}>🔗 Link</button>
      </div>
      <div
        ref={editorRef}
        contentEditable
        onInput={handleInput}
        data-placeholder={placeholder}
        className="rich-text-editor"
        style={{
          ...inputStyle,
          minHeight: 160,
          overflowY: "auto",
          lineHeight: 1.6,
        }}
      />
    </div>
  );
}

function FieldRow({ label, children }) {
  return (
    <div className="field-row" style={{ display: "flex", alignItems: "center", gap: 14, marginBottom: 12, flexWrap: "wrap" }}>
      <label className="field-row-label" style={{ width: 190, fontSize: 13, color: COLORS.inkSoft, flexShrink: 0 }}>{label}</label>
      {children}
    </div>
  );
}

const formCardStyle = {
  background: COLORS.card,
  border: `1px solid ${COLORS.hairline}`,
  borderRadius: 14,
  padding: "24px 26px",
};

const inputStyle = {
  border: `1px solid ${COLORS.hairline}`,
  borderRadius: 11,
  padding: "9px 11px",
  fontSize: 14,
  background: "#fff",
  color: COLORS.ink,
  outline: "none",
  flex: 1,
};

const selectStyle = {
  ...inputStyle,
  background: COLORS.chipBg,
  color: COLORS.ink,
  fontWeight: 500,
  border: `1px solid ${COLORS.seal}`,
  cursor: "pointer",
};

const smallXStyle = {
  background: "none",
  border: "none",
  color: COLORS.seal,
  cursor: "pointer",
  fontSize: 13,
};

/* ================= RADICALS TAB ================= */
function ManagementTab({ userId, isAdmin, tier, lookupCount, lookupLimit, courseName, characterList, wordList, bushouList, decks, onDecksChanged, meaningDisplay, onRequireAuth }) {
  const [subTab, setSubTab] = useState("account"); // account | library

  if (!userId) {
    return (
      <div style={{ textAlign: "center", padding: "60px 20px" }}>
        <div style={{ fontSize: 15, color: COLORS.inkSoft, marginBottom: 16 }}>{t("mgmt_sign_in_required", meaningDisplay)}</div>
        <button type="button" onClick={() => onRequireAuth && onRequireAuth()} className="seal-btn" style={sealBtnStyle}>
          {t("mgmt_sign_in_button", meaningDisplay)}
        </button>
      </div>
    );
  }

  const subTabs = [
    { id: "account", label: t("mgmt_account_tab", meaningDisplay) },
    { id: "library", label: t("mgmt_library_tab", meaningDisplay) },
    { id: "messages", label: t("mgmt_messages_tab", meaningDisplay) },
  ];

  return (
    <div>
      <div style={{ display: "flex", justifyContent: "center", gap: 20, flexWrap: "wrap", borderBottom: `1px solid ${COLORS.hairline}`, marginBottom: 22 }}>
        {subTabs.map((s) => (
          <button
            key={s.id}
            type="button"
            onClick={() => setSubTab(s.id)}
            style={{
              background: "none",
              border: "none",
              borderBottom: `2px solid ${subTab === s.id ? COLORS.seal : "transparent"}`,
              color: subTab === s.id ? COLORS.ink : COLORS.inkSoft,
              fontWeight: subTab === s.id ? 700 : 600,
              fontSize: 14,
              padding: "8px 2px",
              marginBottom: -1,
              cursor: "pointer",
            }}
          >
            {s.label}
          </button>
        ))}
      </div>

      {subTab === "account" ? (
        <AccountManagementTab tier={tier} lookupCount={lookupCount} lookupLimit={lookupLimit} courseName={courseName} meaningDisplay={meaningDisplay} />
      ) : subTab === "messages" ? (
        <UserMessagesTab userId={userId} meaningDisplay={meaningDisplay} />
      ) : (
        <LibraryManagementTab
          userId={userId}
          isAdmin={isAdmin}
          characterList={characterList}
          wordList={wordList}
          bushouList={bushouList}
          decks={decks}
          onDecksChanged={onDecksChanged}
          meaningDisplay={meaningDisplay}
        />
      )}
    </div>
  );
}

function UserMessagesTab({ userId, meaningDisplay }) {
  const [thread, setThread] = useState(null); // null = loading
  const [draft, setDraft] = useState("");
  const [sending, setSending] = useState(false);
  const bottomRef = useRef(null);

  useEffect(() => {
    loadThread();
    // eslint-disable-next-line react-hooks/exhaustive-deps
  }, [userId]);

  useEffect(() => {
    if (thread) bottomRef.current?.scrollIntoView({ block: "nearest" });
  }, [thread]);

  async function loadThread() {
    const { data, error } = await supabase
      .from("messages")
      .select("*")
      .eq("user_id", userId)
      .order("created_at", { ascending: true });
    if (!error) {
      setThread(data || []);
      // Mark the admin's messages as read now that the user is viewing them.
      const unreadIds = (data || []).filter((m) => m.sender === "admin" && !m.read).map((m) => m.id);
      if (unreadIds.length > 0) {
        await supabase.from("messages").update({ read: true }).in("id", unreadIds);
      }
    }
  }

  async function handleSend(e) {
    e.preventDefault();
    if (!draft.trim() || sending) return;
    setSending(true);
    const { error } = await supabase.from("messages").insert({ user_id: userId, sender: "user", message: draft.trim() });
    if (!error) {
      setDraft("");
      await loadThread();
    }
    setSending(false);
  }

  return (
    <div style={{ maxWidth: 480, margin: "0 auto" }}>
      <div
        style={{
          border: `1px solid ${COLORS.hairline}`,
          borderRadius: 12,
          padding: "14px 16px",
          minHeight: 240,
          maxHeight: 420,
          overflowY: "auto",
          marginBottom: 12,
          display: "flex",
          flexDirection: "column",
          gap: 10,
        }}
      >
        {thread === null ? (
          <div style={{ textAlign: "center", color: COLORS.inkSoft, fontSize: 13 }}>{t("loading", meaningDisplay)}</div>
        ) : thread.length === 0 ? (
          <div style={{ textAlign: "center", color: COLORS.inkSoft, fontSize: 13 }}>{t("msg_empty_state", meaningDisplay)}</div>
        ) : (
          thread.map((m) => (
            <div key={m.id} style={{ alignSelf: m.sender === "user" ? "flex-end" : "flex-start", maxWidth: "80%" }}>
              <div style={{ fontSize: 10, color: COLORS.metadata, marginBottom: 2, textAlign: m.sender === "user" ? "right" : "left" }}>
                {m.sender === "user" ? t("msg_from_you", meaningDisplay) : t("msg_from_admin", meaningDisplay)} · {new Date(m.created_at).toLocaleString()}
              </div>
              <div
                style={{
                  background: m.sender === "user" ? COLORS.seal : COLORS.chipBg,
                  color: m.sender === "user" ? "#FBF9EF" : COLORS.ink,
                  borderRadius: 10,
                  padding: "8px 12px",
                  fontSize: 13,
                  lineHeight: 1.5,
                  whiteSpace: "pre-wrap",
                }}
              >
                {m.message}
              </div>
            </div>
          ))
        )}
        <div ref={bottomRef} />
      </div>
      <form onSubmit={handleSend} style={{ display: "flex", gap: 8 }}>
        <input
          value={draft}
          onChange={(e) => setDraft(e.target.value)}
          placeholder={t("msg_placeholder", meaningDisplay)}
          style={{ ...inputStyle, flex: 1, boxSizing: "border-box" }}
        />
        <button type="submit" className="seal-btn" style={{ ...sealBtnStyle, padding: "8px 18px" }} disabled={sending}>
          {t("msg_send", meaningDisplay)}
        </button>
      </form>
    </div>
  );
}

function AccountManagementTab({ tier, lookupCount, lookupLimit, courseName, meaningDisplay }) {
  const tiers = ["Free", "Premium"];
  const currentIndex = tiers.indexOf(tier || "Free");
  const higherTiers = tiers.slice(currentIndex + 1);

  return (
    <div style={{ maxWidth: 480, margin: "0 auto" }}>
      <div style={{ background: COLORS.card, border: `1px solid ${COLORS.hairline}`, borderRadius: 14, padding: "20px 22px", marginBottom: 20 }}>
        <div style={{ fontSize: 11, color: COLORS.metadata, textTransform: "uppercase", letterSpacing: 0.6, marginBottom: 4 }}>
          {t("mgmt_tier_label", meaningDisplay)}
        </div>
        <div style={{ fontSize: 22, fontWeight: 700, color: COLORS.seal, marginBottom: 12 }}>{tier || "Free"}</div>
        <div style={{ fontSize: 13, color: COLORS.inkSoft, marginBottom: courseName ? 4 : 0 }}>
          {t("mgmt_lookup_usage", meaningDisplay, lookupCount ?? 0, lookupLimit ?? 100)}
        </div>
        {courseName && (
          <div style={{ fontSize: 13, color: COLORS.inkSoft }}>{t("mgmt_course_label", meaningDisplay)} {courseName}</div>
        )}
      </div>

      {higherTiers.length > 0 && (
        <div>
          <div style={{ fontSize: 15, fontWeight: 700, color: COLORS.ink, marginBottom: 8, textAlign: "center" }}>
            {t("mgmt_upgrade_title", meaningDisplay)}
          </div>
          <div style={{ fontSize: 13, color: COLORS.inkSoft, textAlign: "center", marginBottom: 16, lineHeight: 1.6 }}>
            {t("mgmt_upgrade_body", meaningDisplay)}
          </div>
          <div style={{ textAlign: "center" }}>
            <a
              href="mailto:hello@minouq.com?subject=Upgrade%20request"
              className="seal-btn"
              style={{ ...sealBtnStyle, textDecoration: "none", display: "inline-block" }}
            >
              {t("mgmt_upgrade_button", meaningDisplay)}
            </a>
          </div>
        </div>
      )}
    </div>
  );
}

function LibraryManagementTab({ userId, isAdmin, characterList, wordList, bushouList, decks, onDecksChanged, meaningDisplay }) {
  const myDecks = (decks || []).filter((d) => d.userId === userId);

  // Personal deck editor -- same shape as the admin deck editor, but every
  // deck this creates is tagged with user_id, and only decks this user owns
  // ever show up here (myDecks above, filtered client-side; RLS also
  // enforces this server-side regardless).
  const [editingDeckId, setEditingDeckId] = useState(null);
  const [deckName, setDeckName] = useState("");
  const [deckLists, setDeckLists] = useState([]);
  const [deckAddType, setDeckAddType] = useState("char");
  const [deckAddList, setDeckAddList] = useState("");
  const [deckMessage, setDeckMessage] = useState(null);

  const availableListsByType = {
    char: Array.from(new Set((characterList || []).flatMap((c) => getLists(c)))).sort((a, b) => a.localeCompare(b, "vi")),
    word: Array.from(new Set((wordList || []).flatMap((w) => w.lists || []))).sort((a, b) => a.localeCompare(b, "vi")),
    bushou: Array.from(new Set((bushouList || []).flatMap((b) => b.lists || []))).sort((a, b) => {
      const numA = parseInt(a, 10);
      const numB = parseInt(b, 10);
      if (!isNaN(numA) && !isNaN(numB)) return numA - numB;
      return a.localeCompare(b, "vi");
    }),
  };

  function startNewDeck() {
    setEditingDeckId("new");
    setDeckName("");
    setDeckLists([]);
    setDeckAddType("char");
    setDeckAddList("");
    setDeckMessage(null);
  }

  function startEditDeck(deck) {
    setEditingDeckId(deck.id);
    setDeckName(deck.name);
    setDeckLists(deck.lists || []);
    setDeckAddType("char");
    setDeckAddList("");
    setDeckMessage(null);
  }

  function addDeckList() {
    if (!deckAddList) return;
    if (deckLists.some((l) => l.content_type === deckAddType && l.list_name === deckAddList)) return;
    setDeckLists([...deckLists, { content_type: deckAddType, list_name: deckAddList }]);
    setDeckAddList("");
  }

  function removeDeckList(contentType, listName) {
    setDeckLists(deckLists.filter((l) => !(l.content_type === contentType && l.list_name === listName)));
  }

  async function saveDeck() {
    if (!deckName.trim()) {
      setDeckMessage({ type: "error", text: t("admin_deck_need_name", meaningDisplay) });
      return;
    }
    let deckId = editingDeckId;
    if (editingDeckId === "new") {
      const { data, error } = await supabase.from("decks").insert({ name: deckName.trim(), user_id: userId }).select().single();
      if (error) {
        setDeckMessage({ type: "error", text: error.message });
        return;
      }
      deckId = data.id;
    } else {
      const { error } = await supabase.from("decks").update({ name: deckName.trim() }).eq("id", editingDeckId);
      if (error) {
        setDeckMessage({ type: "error", text: error.message });
        return;
      }
      await supabase.from("deck_lists").delete().eq("deck_id", editingDeckId);
    }
    if (deckLists.length > 0) {
      const { error } = await supabase
        .from("deck_lists")
        .insert(deckLists.map((l) => ({ deck_id: deckId, content_type: l.content_type, list_name: l.list_name })));
      if (error) {
        setDeckMessage({ type: "error", text: error.message });
        return;
      }
    }
    setEditingDeckId(null);
    if (onDecksChanged) onDecksChanged();
  }

  async function deleteDeck(id) {
    if (!window.confirm(t("admin_deck_confirm_delete", meaningDisplay))) return;
    const { error } = await supabase.from("decks").delete().eq("id", id);
    if (!error && onDecksChanged) onDecksChanged();
  }

  async function makeDeckPublic(id) {
    if (!window.confirm(t("mgmt_confirm_make_public", meaningDisplay))) return;
    const { error } = await supabase.from("decks").update({ user_id: null }).eq("id", id);
    if (!error && onDecksChanged) onDecksChanged();
  }

  // Personal lists -- derived from this user's own custom_* rows only,
  // fetched fresh here rather than from the merged characterList/wordList/
  // bushouList props, so official/shared items never appear as "mine".
  const [myLists, setMyLists] = useState(null); // null = loading
  const CUSTOM_TABLES = [
    { table: "custom_characters", key: "char" },
    { table: "custom_words", key: "word" },
    { table: "custom_bushou", key: "char" },
  ];

  useEffect(() => {
    loadMyLists();
    // eslint-disable-next-line react-hooks/exhaustive-deps
  }, [userId]);

  async function loadMyLists() {
    setMyLists(null);
    const results = await Promise.all(
      CUSTOM_TABLES.map(({ table, key }) => supabase.from(table).select(`${key}, lists`).eq("user_id", userId))
    );
    const counts = new Map();
    results.forEach((res) => {
      (res.data || []).forEach((row) => {
        (row.lists || []).forEach((l) => counts.set(l, (counts.get(l) || 0) + 1));
      });
    });
    setMyLists(
      Array.from(counts.entries())
        .map(([name, count]) => ({ name, count }))
        .sort((a, b) => a.name.localeCompare(b.name, "vi"))
    );
  }

  async function bulkRenameList(oldName, newName) {
    for (const { table, key } of CUSTOM_TABLES) {
      const { data, error } = await supabase.from(table).select(`${key}, lists`).eq("user_id", userId).contains("lists", [oldName]);
      if (error || !data) continue;
      for (const row of data) {
        const newLists = row.lists.map((l) => (l === oldName ? newName : l));
        await supabase.from(table).update({ lists: newLists }).eq("user_id", userId).eq(key, row[key]);
      }
    }
    await loadMyLists();
  }

  async function bulkDeleteList(name) {
    for (const { table, key } of CUSTOM_TABLES) {
      const { data, error } = await supabase.from(table).select(`${key}, lists`).eq("user_id", userId).contains("lists", [name]);
      if (error || !data) continue;
      for (const row of data) {
        const newLists = row.lists.filter((l) => l !== name);
        await supabase.from(table).update({ lists: newLists }).eq("user_id", userId).eq(key, row[key]);
      }
    }
    await loadMyLists();
  }

  function handleRenameList(name) {
    const next = window.prompt(t("mgmt_rename_prompt", meaningDisplay), name);
    if (!next || !next.trim() || next.trim() === name) return;
    bulkRenameList(name, next.trim());
  }

  function handleDeleteList(name) {
    if (!window.confirm(t("mgmt_confirm_delete_list", meaningDisplay, name))) return;
    bulkDeleteList(name);
  }

  return (
    <div style={{ maxWidth: 560, margin: "0 auto" }}>
      {/* Personal decks */}
      <div style={{ display: "flex", justifyContent: "space-between", alignItems: "center", marginBottom: 12 }}>
        <div style={{ fontSize: 12.5, fontWeight: 600, color: COLORS.gold, textTransform: "uppercase", letterSpacing: 0.8 }}>
          {t("mgmt_my_decks_title", meaningDisplay)}
        </div>
        <button type="button" onClick={startNewDeck} className="seal-btn" style={{ ...sealBtnStyle, padding: "6px 14px", fontSize: 12 }}>
          {t("admin_deck_new", meaningDisplay)}
        </button>
      </div>

      {editingDeckId && (
        <div style={{ background: COLORS.card, border: `1px solid ${COLORS.hairline}`, borderRadius: 11, padding: "14px 16px", marginBottom: 14 }}>
          <input
            value={deckName}
            onChange={(e) => setDeckName(e.target.value)}
            placeholder={t("admin_deck_name_placeholder", meaningDisplay)}
            style={{ ...inputStyle, width: "100%", boxSizing: "border-box", marginBottom: 10, fontWeight: 600 }}
          />

          <div style={{ fontSize: 11.5, color: COLORS.inkSoft, marginBottom: 6 }}>{t("admin_deck_included_lists", meaningDisplay)}</div>
          {deckLists.length === 0 ? (
            <div style={{ fontSize: 12, color: COLORS.metadata, marginBottom: 10 }}>{t("admin_deck_none_included", meaningDisplay)}</div>
          ) : (
            <div style={{ display: "flex", flexWrap: "wrap", gap: 6, marginBottom: 10 }}>
              {deckLists.map((l) => (
                <span
                  key={`${l.content_type}:${l.list_name}`}
                  style={{ display: "inline-flex", alignItems: "center", gap: 5, background: COLORS.chipBg, border: `1px solid ${COLORS.hairline}`, borderRadius: 999, padding: "3px 9px", fontSize: 11.5 }}
                >
                  <span style={{ color: COLORS.seal, fontWeight: 600 }}>
                    {l.content_type === "char" ? t("admin_deck_type_char", meaningDisplay) : l.content_type === "word" ? t("admin_deck_type_word", meaningDisplay) : t("admin_deck_type_bushou", meaningDisplay)}
                  </span>
                  · {l.content_type === "bushou" ? displayListName(l.list_name, meaningDisplay) : l.list_name}
                  <button
                    type="button"
                    onClick={() => removeDeckList(l.content_type, l.list_name)}
                    style={{ background: "none", border: "none", color: COLORS.error, cursor: "pointer", padding: 0, fontSize: 12, lineHeight: 1 }}
                  >
                    ✕
                  </button>
                </span>
              ))}
            </div>
          )}

          <div style={{ display: "flex", gap: 6, marginBottom: 12, flexWrap: "wrap" }}>
            <select value={deckAddType} onChange={(e) => { setDeckAddType(e.target.value); setDeckAddList(""); }} style={{ ...selectStyle, width: 130 }}>
              <option value="char" style={{ background: COLORS.chipBg }}>{t("admin_deck_type_char", meaningDisplay)}</option>
              <option value="word" style={{ background: COLORS.chipBg }}>{t("admin_deck_type_word", meaningDisplay)}</option>
              <option value="bushou" style={{ background: COLORS.chipBg }}>{t("admin_deck_type_bushou", meaningDisplay)}</option>
            </select>
            <select value={deckAddList} onChange={(e) => setDeckAddList(e.target.value)} style={{ ...selectStyle, width: 180 }}>
              <option value="" style={{ background: COLORS.chipBg }}>{t("admin_deck_choose_list", meaningDisplay)}</option>
              {(availableListsByType[deckAddType] || []).map((l) => (
                <option key={l} value={l} style={{ background: COLORS.chipBg }}>
                  {deckAddType === "bushou" ? displayListName(l, meaningDisplay) : l}
                </option>
              ))}
            </select>
            <button type="button" onClick={addDeckList} className="ghost-btn" style={{ ...ghostBtnStyle, padding: "6px 12px", fontSize: 12 }}>
              {t("admin_deck_add", meaningDisplay)}
            </button>
          </div>

          {deckMessage && (
            <div style={{ fontSize: 12, fontWeight: 600, color: deckMessage.type === "error" ? COLORS.error : COLORS.seal, marginBottom: 10 }}>
              {deckMessage.text}
            </div>
          )}
          <div style={{ display: "flex", gap: 8 }}>
            <button type="button" onClick={saveDeck} className="seal-btn" style={{ ...sealBtnStyle, padding: "8px 16px", fontSize: 13 }}>
              {t("admin_deck_save", meaningDisplay)}
            </button>
            <button type="button" onClick={() => setEditingDeckId(null)} className="ghost-btn" style={{ ...ghostBtnStyle, padding: "8px 16px", fontSize: 13 }}>
              {t("admin_deck_cancel", meaningDisplay)}
            </button>
          </div>
        </div>
      )}

      {myDecks.length === 0 ? (
        <div style={{ textAlign: "center", color: COLORS.inkSoft, padding: 16, fontSize: 13 }}>{t("admin_deck_none", meaningDisplay)}</div>
      ) : (
        <div style={{ display: "flex", flexDirection: "column", gap: 8, marginBottom: 28 }}>
          {myDecks.map((d) => (
            <div key={d.id} style={{ background: COLORS.card, border: `1px solid ${COLORS.hairline}`, borderRadius: 11, padding: "12px 16px", display: "flex", justifyContent: "space-between", alignItems: "center", gap: 10, flexWrap: "wrap" }}>
              <div>
                <div style={{ fontSize: 13.5, fontWeight: 600, color: COLORS.ink }}>📦 {d.name}</div>
                <div style={{ fontSize: 11, color: COLORS.metadata }}>{t("admin_deck_list_count", meaningDisplay, (d.lists || []).length)}</div>
              </div>
              <div style={{ display: "flex", gap: 6 }}>
                {isAdmin && (
                  <button
                    type="button"
                    onClick={() => makeDeckPublic(d.id)}
                    className="ghost-btn"
                    style={{ ...ghostBtnStyle, padding: "5px 10px", fontSize: 11.5, borderColor: COLORS.gold, color: COLORS.gold }}
                  >
                    {t("mgmt_make_public", meaningDisplay)}
                  </button>
                )}
                <button type="button" onClick={() => startEditDeck(d)} className="ghost-btn" style={{ ...ghostBtnStyle, padding: "5px 10px", fontSize: 11.5 }}>
                  {t("admin_deck_edit", meaningDisplay)}
                </button>
                <button
                  type="button"
                  onClick={() => deleteDeck(d.id)}
                  className="ghost-btn"
                  style={{ ...ghostBtnStyle, padding: "5px 10px", fontSize: 11.5, borderColor: COLORS.error, color: COLORS.error }}
                >
                  {t("admin_deck_delete", meaningDisplay)}
                </button>
              </div>
            </div>
          ))}
        </div>
      )}

      {/* Personal lists */}
      <div style={{ fontSize: 12.5, fontWeight: 600, color: COLORS.gold, textTransform: "uppercase", letterSpacing: 0.8, marginBottom: 8 }}>
        {t("mgmt_my_lists_title", meaningDisplay)}
      </div>
      <div style={{ fontSize: 12, color: COLORS.inkSoft, marginBottom: 14, lineHeight: 1.5 }}>{t("mgmt_lists_description", meaningDisplay)}</div>

      {myLists === null ? (
        <div style={{ textAlign: "center", color: COLORS.inkSoft, padding: 16 }}>{t("loading", meaningDisplay)}</div>
      ) : myLists.length === 0 ? (
        <div style={{ textAlign: "center", color: COLORS.inkSoft, padding: 16, fontSize: 13 }}>{t("mgmt_no_personal_lists", meaningDisplay)}</div>
      ) : (
        <div style={{ display: "flex", flexDirection: "column", gap: 8 }}>
          {myLists.map((l) => (
            <div key={l.name} style={{ background: COLORS.card, border: `1px solid ${COLORS.hairline}`, borderRadius: 11, padding: "10px 16px", display: "flex", justifyContent: "space-between", alignItems: "center", gap: 10, flexWrap: "wrap" }}>
              <div>
                <div style={{ fontSize: 13.5, fontWeight: 600, color: COLORS.ink }}>{l.name}</div>
                <div style={{ fontSize: 11, color: COLORS.metadata }}>{t("mgmt_item_count", meaningDisplay, l.count)}</div>
              </div>
              <div style={{ display: "flex", gap: 6 }}>
                <button type="button" onClick={() => handleRenameList(l.name)} className="ghost-btn" style={{ ...ghostBtnStyle, padding: "5px 10px", fontSize: 11.5 }}>
                  {t("mgmt_rename_list", meaningDisplay)}
                </button>
                <button
                  type="button"
                  onClick={() => handleDeleteList(l.name)}
                  className="ghost-btn"
                  style={{ ...ghostBtnStyle, padding: "5px 10px", fontSize: 11.5, borderColor: COLORS.error, color: COLORS.error }}
                >
                  {t("mgmt_delete_list", meaningDisplay)}
                </button>
              </div>
            </div>
          ))}
        </div>
      )}
    </div>
  );
}

function LibraryTab(props) {
  const {
    bushouList, characterList, wordList, meaningDisplay,
    onAddBushou, officialBushouKeys, overrideBushouKeys, onPromoteBushou, onWithdrawBushou,
    onDeleteCharacter, onDeleteCharacterFromOfficial, onUpdateCharacter, officialCharKeys, overrideCharKeys, onPromoteCharacter, onWithdrawCharacter,
    findBushou, onAddWord, onDeleteWord, onDeleteWordFromOfficial, officialWordKeys, overrideWordKeys, onPromoteWord, onWithdrawWord,
    isAdmin, checkListAccess, onViewPremium, userId,
    customWords, onAddCharacter, onRequireAuth, onQuotaUpdate,
    pendingSubTab, onConsumePendingSubTab,
  } = props;
  const [subTab, setSubTab] = useState("radicals"); // radicals | hanzi | vocab | add

  useEffect(() => {
    if (pendingSubTab) {
      setSubTab(pendingSubTab);
      if (onConsumePendingSubTab) onConsumePendingSubTab();
    }
  }, [pendingSubTab, onConsumePendingSubTab]);

  const subTabs = [
    { id: "radicals", label: t("tab_radicals", meaningDisplay) },
    { id: "hanzi", label: t("tab_hanzi", meaningDisplay) },
    { id: "vocab", label: t("tab_vocab", meaningDisplay) },
  ];

  return (
    <div>
      <div style={{ display: "flex", justifyContent: "center", alignItems: "center", gap: 20, flexWrap: "wrap", borderBottom: `1px solid ${COLORS.hairline}`, marginBottom: 22 }}>
        {subTabs.map((s) => (
          <button
            key={s.id}
            type="button"
            onClick={() => setSubTab(s.id)}
            style={{
              background: "none",
              border: "none",
              borderBottom: `2px solid ${subTab === s.id ? COLORS.seal : "transparent"}`,
              color: subTab === s.id ? COLORS.ink : COLORS.inkSoft,
              fontWeight: subTab === s.id ? 700 : 600,
              fontSize: 14,
              padding: "8px 2px",
              marginBottom: -1,
              cursor: "pointer",
            }}
          >
            {s.label}
          </button>
        ))}
        <button
          type="button"
          onClick={() => setSubTab("add")}
          style={{
            border: `1.5px solid ${COLORS.gold}`,
            borderRadius: 999,
            background: subTab === "add" ? COLORS.gold : "transparent",
            color: subTab === "add" ? "#FBF9EF" : COLORS.gold,
            fontWeight: 700,
            fontSize: 13,
            padding: "6px 14px",
            marginBottom: 6,
            cursor: "pointer",
          }}
        >
          + {t("tab_add", meaningDisplay)}
        </button>
      </div>

      {subTab === "add" ? (
        <AddTab
          bushouList={bushouList}
          characterList={characterList}
          wordList={wordList}
          customWords={customWords}
          onAddCharacter={onAddCharacter}
          onAddBushou={onAddBushou}
          onUpdateCharacter={onUpdateCharacter}
          onDeleteCharacter={onDeleteCharacter}
          onAddWord={onAddWord}
          onDeleteWord={onDeleteWord}
          userId={userId}
          onRequireAuth={onRequireAuth}
          onViewPremium={onViewPremium}
          onQuotaUpdate={onQuotaUpdate}
          meaningDisplay={meaningDisplay}
        />
      ) : subTab === "radicals" ? (
        <RadicalsTab
          bushouList={bushouList}
          onAddBushou={onAddBushou}
          isAdmin={isAdmin}
          officialBushouKeys={officialBushouKeys}
          overrideBushouKeys={overrideBushouKeys}
          onPromoteBushou={onPromoteBushou}
          onWithdrawBushou={onWithdrawBushou}
          meaningDisplay={meaningDisplay}
        />
      ) : subTab === "hanzi" ? (
        <CharacterListPanel
          characterList={characterList}
          bushouList={bushouList}
          onDeleteCharacter={onDeleteCharacter}
          onDeleteCharacterFromOfficial={onDeleteCharacterFromOfficial}
          onUpdateCharacter={onUpdateCharacter}
          onAddBushou={onAddBushou}
          isAdmin={isAdmin}
          officialCharKeys={officialCharKeys}
          overrideCharKeys={overrideCharKeys}
          onPromoteCharacter={onPromoteCharacter}
          onWithdrawCharacter={onWithdrawCharacter}
          checkListAccess={checkListAccess}
          onViewPremium={onViewPremium}
          meaningDisplay={meaningDisplay}
          userId={userId}
        />
      ) : (
        <WordListPanel
          wordList={wordList}
          characterList={characterList}
          findBushou={findBushou}
          onAddWord={onAddWord}
          onDeleteWord={onDeleteWord}
          onDeleteWordFromOfficial={onDeleteWordFromOfficial}
          isAdmin={isAdmin}
          officialWordKeys={officialWordKeys}
          overrideWordKeys={overrideWordKeys}
          onPromoteWord={onPromoteWord}
          onWithdrawWord={onWithdrawWord}
          checkListAccess={checkListAccess}
          onViewPremium={onViewPremium}
          meaningDisplay={meaningDisplay}
          userId={userId}
        />
      )}
    </div>
  );
}

function RadicalsTab({ bushouList, onAddBushou, isAdmin, officialBushouKeys, overrideBushouKeys, onPromoteBushou, onWithdrawBushou, meaningDisplay }) {
  const [query, setQuery] = useState("");
  const [listFilter, setListFilter] = useState("Tất cả");
  const allLists = useMemo(() => {
    const set = new Set();
    bushouList.forEach((b) => (b.lists || []).forEach((l) => set.add(l)));
    // Numeric-first sort so "10 nét" / "11-17 nét" land after "9 nét"
    // instead of a plain alphabetical sort putting "10" right after "1".
    return Array.from(set).sort((a, b) => {
      const numA = parseInt(a, 10);
      const numB = parseInt(b, 10);
      if (!isNaN(numA) && !isNaN(numB)) return numA - numB;
      if (!isNaN(numA)) return -1;
      if (!isNaN(numB)) return 1;
      return a.localeCompare(b);
    });
  }, [bushouList]);

  const filtered = bushouList.filter((b) => {
    if (listFilter !== "Tất cả" && !(b.lists || []).includes(listFilter)) return false;
    const q = query.trim().toLowerCase();
    if (!q) return true;
    return (
      b.char.includes(q) ||
      b.pinyin.toLowerCase().includes(q) ||
      b.meaning.toLowerCase().includes(q) ||
      b.sv.toLowerCase().includes(q)
    );
  });

  // Group by stroke count, ascending — the traditional 部首 chart order.
  // Radicals without a known stroke count (user-added via the "add" tab)
  // are bucketed under "?" and shown last.
  const groups = useMemo(() => {
    const map = new Map();
    filtered
      .slice()
      .sort((a, b) => {
        const sa = typeof a.strokes === "number" ? a.strokes : 999;
        const sb = typeof b.strokes === "number" ? b.strokes : 999;
        if (sa !== sb) return sa - sb;
        return a.char.localeCompare(b.char, "zh");
      })
      .forEach((b) => {
        const key = typeof b.strokes === "number" ? b.strokes : "?";
        if (!map.has(key)) map.set(key, []);
        map.get(key).push(b);
      });
    return Array.from(map.entries());
  }, [filtered]);

  return (
    <div>
      <div
        style={{
          background: COLORS.card,
          border: `1px solid ${COLORS.hairline}`,
          borderRadius: 14,
          padding: "16px 18px",
          marginBottom: 20,
          fontSize: 13,
          color: COLORS.inkSoft,
          lineHeight: 1.7,
        }}
      >
        <p style={{ margin: 0, marginBottom: 10 }}>
          {t("radicals_header_p1", meaningDisplay)}
        </p>
        <p style={{ margin: 0, marginBottom: 10 }}>
          {t("radicals_header_p2", meaningDisplay)}
        </p>
        <p style={{ margin: 0, fontStyle: "italic" }}>
          {t("radicals_header_p3", meaningDisplay)}
        </p>
      </div>

      <div style={{ textAlign: "center", marginBottom: 16, display: "flex", justifyContent: "center", gap: 8, flexWrap: "wrap" }}>
        <input
          value={query}
          onChange={(e) => setQuery(e.target.value)}
          placeholder={t("radicals_search_placeholder", meaningDisplay)}
          style={{ ...inputStyle, width: 280, textAlign: "center" }}
        />
        <select value={listFilter} onChange={(e) => setListFilter(e.target.value)} style={{ ...selectStyle, width: 170 }}>
          <option value="Tất cả" style={{ background: COLORS.chipBg, color: COLORS.ink, fontWeight: 700 }}>{t("play_all_lists", meaningDisplay)}</option>
          {allLists.map((l) => (
            <option key={l} value={l} style={{ background: COLORS.chipBg, color: COLORS.ink, fontWeight: 700 }}>
              {displayListName(l, meaningDisplay)}
            </option>
          ))}
        </select>
      </div>
      <div style={{ fontSize: 12.5, color: COLORS.inkSoft, textAlign: "center", marginBottom: 20 }}>
        {t("radicals_count", meaningDisplay, filtered.length, bushouList.length)}
      </div>

      {groups.map(([strokeCount, items]) => (
        <div key={strokeCount} style={{ marginBottom: 22 }}>
          <div
            style={{
              display: "flex",
              alignItems: "center",
              gap: 10,
              marginBottom: 10,
            }}
          >
            <div
              style={{
                width: 30,
                height: 30,
                borderRadius: "50%",
                background: COLORS.seal,
                color: "#FBF9EF",
                display: "flex",
                alignItems: "center",
                justifyContent: "center",
                fontSize: 13,
                fontWeight: 700,
                flexShrink: 0,
              }}
            >
              {strokeCount}
            </div>
            <div style={{ fontSize: 12.5, color: COLORS.inkSoft, letterSpacing: 0.4 }}>
              {strokeCount === "?" ? t("radicals_stroke_unknown", meaningDisplay) : t("radicals_stroke_count", meaningDisplay, strokeCount)}
            </div>
            <div style={{ flex: 1, borderBottom: `1px dashed ${COLORS.grid}` }} />
          </div>

          <div style={{ display: "grid", gridTemplateColumns: "repeat(auto-fill, minmax(150px, 1fr))", gap: 12 }}>
            {items.map((b) => (
              <RadicalCard
                key={b.char}
                b={b}
                onAddBushou={onAddBushou}
                isAdmin={isAdmin}
                isOfficial={officialBushouKeys ? officialBushouKeys.has(b.char) : false}
                hasOverride={overrideBushouKeys ? overrideBushouKeys.has(b.char) : false}
                onPromoteBushou={onPromoteBushou}
                onWithdrawBushou={onWithdrawBushou}
                meaningDisplay={meaningDisplay}
                allBushouLists={allLists}
              />
            ))}
          </div>
        </div>
      ))}
    </div>
  );
}

/* ---------- A single radical card: view mode + edit mode.
   Saving just re-upserts the same char via onAddBushou (addBushouRow),
   which already overwrites on conflict — so "add" and "edit" are the same
   operation under the hood, exactly like character editing works. ---------- */
function RadicalCard({ b, onAddBushou, isAdmin, isOfficial, hasOverride, onPromoteBushou, onWithdrawBushou, meaningDisplay, allBushouLists }) {
  const [mode, setMode] = useState("view"); // view | edit
  const [strokeOrderOpen, setStrokeOrderOpen] = useState(false);
  const [pinyin, setPinyin] = useState(b.pinyin);
  const [meaning, setMeaning] = useState(b.meaning);
  const [sv, setSv] = useState(b.sv);
  const [strokes, setStrokes] = useState(typeof b.strokes === "number" ? String(b.strokes) : "");
  const [lists, setLists] = useState(b.lists || []);
  const [listTypeahead, setListTypeahead] = useState("");
  const [defaultStatus, setDefaultStatus] = useState("idle"); // idle | working | error

  // A personal edit sitting on top of the official value always means
  // "publish this" (promote) — checking hasOverride FIRST is what makes
  // editing an already-official item push the fix instead of accidentally
  // withdrawing it.
  async function handleToggleDefault() {
    setDefaultStatus("working");
    try {
      if (hasOverride || !isOfficial) await onPromoteBushou(b);
      else await onWithdrawBushou(b);
      setDefaultStatus("idle");
    } catch (e) {
      setDefaultStatus("error");
      setTimeout(() => setDefaultStatus("idle"), 2500);
    }
  }

  function startEdit() {
    setPinyin(b.pinyin);
    setMeaning(b.meaning);
    setSv(b.sv);
    setStrokes(typeof b.strokes === "number" ? String(b.strokes) : "");
    setLists(b.lists || []);
    setListTypeahead("");
    setMode("edit");
  }

  function addList(name) {
    const trimmed = name.trim();
    if (trimmed && !lists.includes(trimmed)) setLists([...lists, trimmed]);
    setListTypeahead("");
  }

  function saveEdit() {
    if (!pinyin.trim() || !meaning.trim() || !sv.trim()) return;
    const strokesNum = parseInt(strokes, 10);
    onAddBushou &&
      onAddBushou({
        char: b.char,
        pinyin: pinyin.trim(),
        meaning: meaning.trim(),
        sv: sv.trim(),
        strokes: Number.isFinite(strokesNum) && strokesNum > 0 ? strokesNum : undefined,
        lists,
      });
    setMode("view");
  }

  return (
    <div
      style={{
        background: COLORS.card,
        border: `1px solid ${COLORS.hairline}`,
        borderRadius: 11,
        padding: "12px 10px",
        textAlign: "center",
        position: "relative",
      }}
    >
      {mode === "view" && (
        <button
          type="button"
          onClick={() => setStrokeOrderOpen(true)}
          title={t("radicals_view_stroke_order", meaningDisplay)}
          style={{
            position: "absolute",
            top: 6,
            left: 6,
            width: 20,
            height: 20,
            lineHeight: "18px",
            padding: 0,
            fontSize: 11,
            border: `1px solid ${COLORS.hairline}`,
            borderRadius: "50%",
            background: COLORS.chipBg,
            color: COLORS.sealDark,
            cursor: "pointer",
          }}
        >
          ✍️
        </button>
      )}

      {mode === "view" && (isAdmin || hasOverride || !isOfficial) && (
        <button
          type="button"
          onClick={startEdit}
          title="Sửa bộ thủ này"
          style={{
            position: "absolute",
            top: 6,
            right: 6,
            width: 20,
            height: 20,
            lineHeight: "18px",
            padding: 0,
            fontSize: 11,
            border: `1px solid ${COLORS.hairline}`,
            borderRadius: "50%",
            background: COLORS.chipBg,
            color: COLORS.gold,
            cursor: "pointer",
          }}
        >
          ✎
        </button>
      )}

      {mode === "edit" ? (
        <div style={{ textAlign: "left" }}>
          <div style={{ fontFamily: "'Noto Serif SC', 'STKaiti', 'Kaiti SC', serif", fontSize: 26, color: COLORS.ink, textAlign: "center", marginBottom: 8 }}>
            {b.char}
          </div>
          <label style={{ fontSize: 10, color: COLORS.inkSoft, display: "block", marginBottom: 2 }}>Pinyin</label>
          <input value={pinyin} onChange={(e) => setPinyin(e.target.value)} style={{ ...inputStyle, width: "100%", marginBottom: 6, fontSize: 12.5 }} />
          <label style={{ fontSize: 10, color: COLORS.inkSoft, display: "block", marginBottom: 2 }}>Nghĩa</label>
          <input value={meaning} onChange={(e) => setMeaning(e.target.value)} style={{ ...inputStyle, width: "100%", marginBottom: 6, fontSize: 12.5 }} />
          <label style={{ fontSize: 10, color: COLORS.inkSoft, display: "block", marginBottom: 2 }}>Hán Việt</label>
          <input value={sv} onChange={(e) => setSv(e.target.value)} style={{ ...inputStyle, width: "100%", marginBottom: 6, fontSize: 12.5 }} />
          <label style={{ fontSize: 10, color: COLORS.inkSoft, display: "block", marginBottom: 2 }}>Số nét</label>
          <input
            value={strokes}
            onChange={(e) => setStrokes(e.target.value.replace(/[^0-9]/g, ""))}
            style={{ ...inputStyle, width: "100%", marginBottom: 8, fontSize: 12.5 }}
          />
          <label style={{ fontSize: 10, color: COLORS.inkSoft, display: "block", marginBottom: 2 }}>{t("radicals_lists_field_label", meaningDisplay)}</label>
          {lists.length > 0 && (
            <div style={{ display: "flex", flexWrap: "wrap", gap: 4, marginBottom: 6 }}>
              {lists.map((l) => (
                <span
                  key={l}
                  style={{ display: "inline-flex", alignItems: "center", gap: 4, background: COLORS.chipBg, border: `1px solid ${COLORS.hairline}`, borderRadius: 999, padding: "2px 8px", fontSize: 11 }}
                >
                  {l}
                  <button
                    type="button"
                    onClick={() => setLists(lists.filter((x) => x !== l))}
                    style={{ background: "none", border: "none", color: COLORS.error, cursor: "pointer", padding: 0, fontSize: 12, lineHeight: 1 }}
                  >
                    ✕
                  </button>
                </span>
              ))}
            </div>
          )}
          <div style={{ display: "flex", gap: 4, marginBottom: 8 }}>
            <input
              value={listTypeahead}
              onChange={(e) => setListTypeahead(e.target.value)}
              onKeyDown={(e) => {
                if (e.key === "Enter") {
                  e.preventDefault();
                  addList(listTypeahead);
                }
              }}
              placeholder="vd: 1 nét… rồi Enter"
              list="existing-bushou-lists"
              style={{ ...inputStyle, fontSize: 11.5, padding: "5px 8px" }}
            />
            <datalist id="existing-bushou-lists">
              {(allBushouLists || []).map((l) => (
                <option key={l} value={l} />
              ))}
            </datalist>
            <button type="button" onClick={() => addList(listTypeahead)} className="ghost-btn" style={{ ...ghostBtnStyle, padding: "5px 10px", fontSize: 11.5 }}>
              +
            </button>
          </div>
          <div style={{ display: "flex", justifyContent: "center", gap: 6 }}>
            <button type="button" onClick={saveEdit} className="seal-btn" style={{ ...sealBtnStyle, padding: "6px 14px", fontSize: 12 }}>
              Lưu
            </button>
            <button type="button" onClick={() => setMode("view")} className="ghost-btn" style={{ ...ghostBtnStyle, padding: "6px 14px", fontSize: 12 }}>
              Hủy
            </button>
          </div>
        </div>
      ) : (
        <>
          <div style={{ fontFamily: "'Noto Serif SC', 'STKaiti', 'Kaiti SC', serif", fontSize: 30, color: COLORS.ink }}>{b.char}</div>
          <div style={{ fontSize: 12.5, color: COLORS.sealDark, marginTop: 4 }}>{b.pinyin}</div>
          {meaningDisplay !== "vi" && (
            <div style={{ fontSize: 11.5, color: COLORS.inkSoft, marginTop: 2 }}>{b.meaning}</div>
          )}
          {meaningDisplay !== "en" && (
            <div style={{ fontSize: 11.5, color: COLORS.bamboo, marginTop: 2, fontWeight: 600 }}>HV: {b.sv}</div>
          )}
          {isAdmin && (
            <button
              type="button"
              onClick={handleToggleDefault}
              disabled={defaultStatus === "working"}
              title={
                hasOverride || !isOfficial
                  ? t("radicals_set_default_tooltip", meaningDisplay)
                  : t("radicals_withdraw_tooltip", meaningDisplay)
              }
              style={{
                marginTop: 6,
                fontSize: 10,
                padding: "3px 8px",
                borderRadius: 999,
                border: `1px solid ${!hasOverride && isOfficial ? COLORS.bamboo : COLORS.gold}`,
                background: !hasOverride && isOfficial ? "rgba(80,133,165,0.12)" : "rgba(80,133,165,0.06)",
                color: defaultStatus === "error" ? COLORS.error : !hasOverride && isOfficial ? COLORS.bamboo : COLORS.gold,
                cursor: defaultStatus === "working" ? "default" : "pointer",
                opacity: defaultStatus === "working" ? 0.6 : 1,
              }}
            >
              {defaultStatus === "working"
                ? t("radicals_working", meaningDisplay)
                : defaultStatus === "error"
                ? t("radicals_error_retry", meaningDisplay)
                : hasOverride && isOfficial
                ? t("radicals_update_default", meaningDisplay)
                : !hasOverride && isOfficial
                ? t("radicals_is_default", meaningDisplay)
                : t("radicals_set_default", meaningDisplay)}
            </button>
          )}
          {isOfficial && (
            <div style={{ marginTop: 6 }}>
              <SuggestRevisionButton contentType="bushou" itemKey={b.char} meaningDisplay={meaningDisplay} />
            </div>
          )}
        </>
      )}

      {strokeOrderOpen && <StrokeOrderModal char={b.char} onClose={() => setStrokeOrderOpen(false)} meaningDisplay={meaningDisplay} />}
    </div>
  );
}
