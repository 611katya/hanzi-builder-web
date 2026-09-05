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
const FONT_IMPORT =
  "@import url('https://fonts.googleapis.com/css2?family=Noto+Serif+SC:wght@400;600;700&family=Cormorant+Garamond:ital,wght@0,500;0,600;0,700;1,600&family=Inter:wght@400;500;600;700&display=swap');";

const COLORS = {
  paper: "#F2F0E1",
  paperDark: "#E7E4CE",
  card: "#FAF9EF",
  ink: "#2A2A16",
  inkSoft: "#7A7A56",
  seal: "#556B2F",
  sealDark: "#3E4E22",
  bamboo: "#595900",
  bambooDark: "#3F3F00",
  gold: "#584C25",
  grid: "#DDD9BB",
  chipBg: "#F6F4E6",
  error: "#A6432E",
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
  tab_play: { vi: "Ghép bộ thủ", en: "Assemble Radicals" },
  tab_flashcards: { vi: "Flashcard", en: "Flashcard" },
  tab_writing: { vi: "✍️ Luyện viết", en: "✍️ Writing Practice" },
  tab_add: { vi: "Tạo thẻ từ mới", en: "Create New Card" },
  tab_radicals: { vi: "Bộ thủ", en: "Radicals" },
  tab_hanzi: { vi: "Hán tự", en: "Characters" },
  tab_vocab: { vi: "Từ vựng", en: "Vocabulary" },
  tab_premium: { vi: "★ Thành viên", en: "★ Membership" },
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
  wp_reveal: { vi: "Hiện chữ đúng", en: "Reveal Correct Character" },
  wp_undo_stroke: { vi: "↩ Xóa nét trước", en: "↩ Undo Last Stroke" },
  wp_recall_erase: { vi: "↩ Tẩy nét cuối", en: "↩ Erase Last Stroke" },
  wp_clear_all: { vi: "🗑 Xóa hết, viết lại", en: "🗑 Clear All, Start Over" },
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
  wp_dots_green: { vi: "xanh", en: "green" },
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

  // Tạo thẻ từ mới (Add tab)
  add_intro: {
    vi: "Nhập một chữ Hán hoàn chỉnh cùng nghĩa, pinyin, âm Hán Việt, và xếp vào một danh sách (list) tuỳ chọn.",
    en: "Enter a complete Chinese character along with its meaning, pinyin, Sino-Vietnamese reading, and an optional list.",
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
  word_hanviet_label: { vi: "Hán Việt (tùy chọn)", en: "Sino-Vietnamese (optional)" },
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

  // Hán tự (Character list panel)
  hanzi_panel_title: { vi: "Danh sách Hán tự trong kho dữ liệu", en: "Character List in Storage" },
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

function uid() {
  return Math.random().toString(36).slice(2, 10);
}

/* ---------- Supabase row <-> app object conversion ---------- */
function rowToBushou(row) {
  return { char: row.char, pinyin: row.pinyin, meaning: row.meaning, sv: row.sv, strokes: row.strokes };
}
function bushouToRow(b, userId) {
  return {
    user_id: userId,
    char: b.char,
    pinyin: b.pinyin,
    meaning: b.meaning,
    sv: b.sv,
    strokes: typeof b.strokes === "number" ? b.strokes : null,
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
    border: `1px solid ${COLORS.grid}`,
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


function Chip({ info, onClick, disabled, big, tone }) {
  return (
    <button
      onClick={onClick}
      disabled={disabled}
      title={info ? `${info.pinyin} · ${info.meaning} · SV: ${info.sv}` : ""}
      className="hanzi-chip"
      style={{
        fontFamily: "KaiTi, 'STKaiti', 'Kaiti SC', 'Noto Serif SC', serif",
        fontSize: big ? 30 : 26,
        width: big ? 56 : 48,
        height: big ? 56 : 48,
        lineHeight: 1,
        display: "flex",
        alignItems: "center",
        justifyContent: "center",
        borderRadius: 8,
        border: `1.5px solid ${tone === "seal" ? COLORS.seal : COLORS.gold}`,
        background: disabled ? "rgba(169,130,47,0.12)" : COLORS.chipBg,
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
function StrokeOrderModal({ char, onClose }) {
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
            border: `1px solid ${COLORS.grid}`,
            borderRadius: "50%",
            background: COLORS.chipBg,
            color: COLORS.inkSoft,
            cursor: "pointer",
          }}
        >
          ✕
        </button>

        <div style={{ fontSize: 11, color: COLORS.inkSoft, textTransform: "uppercase", letterSpacing: 0.6, marginBottom: 10 }}>
          Thứ tự nét bút
        </div>

        <div
          ref={targetRef}
          style={{
            width: 260,
            height: 260,
            margin: "0 auto",
            border: `2px solid ${COLORS.grid}`,
            borderRadius: 6,
            background: COLORS.card,
            display: status === "error" ? "none" : "block",
          }}
        />

        {status === "error" && (
          <div style={{ width: 260, height: 260, margin: "0 auto", display: "flex", alignItems: "center", justifyContent: "center", color: COLORS.inkSoft, fontSize: 13, padding: 16 }}>
            Chưa có dữ liệu nét bút cho chữ "{char}" trong nguồn dữ liệu.
          </div>
        )}

        <button
          type="button"
          onClick={replay}
          disabled={status === "error"}
          className="ghost-btn"
          style={{ ...ghostBtnStyle, marginTop: 14, opacity: status === "error" ? 0.4 : 1 }}
        >
          ▶ Xem lại
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
            fontFamily: "'Inter', sans-serif",
            padding: 30,
            maxWidth: 600,
            margin: "40px auto",
            background: "#FAF9EF",
            border: "1px solid #DDD9BB",
            borderRadius: 10,
            color: "#2A2A16",
          }}
        >
          <div style={{ fontWeight: 700, marginBottom: 8, color: "#A6432E" }}>Đã xảy ra lỗi · Something went wrong</div>
          <div style={{ fontSize: 13, marginBottom: 14, color: "#7A7A56" }}>
            {String(this.state.error && this.state.error.message ? this.state.error.message : this.state.error)}
          </div>
          <button
            onClick={() => this.setState({ error: null })}
            style={{
              background: "#A6432E",
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
  const [tab, setTab] = useState("play");

  // The shared default data, loaded from Supabase for EVERYONE (including
  // guests, via public SELECT policies) so admin corrections go live for
  // every visitor without a code deploy. Falls back to the hardcoded
  // SEED_ arrays if the tables are empty/unreachable, so the app never
  // breaks even if this fetch has a problem.
  const [officialBushou, setOfficialBushou] = useState(null); // null = not loaded yet
  const [officialChars, setOfficialChars] = useState(null);
  const [officialWords, setOfficialWords] = useState(null);
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
        setLookupLimit(data.lookup_limit != null ? data.lookup_limit : 100);
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
        fontFamily: "'Inter', sans-serif",
        background: `radial-gradient(circle at 15% 8%, ${COLORS.paperDark}, ${COLORS.paper} 55%)`,
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
        input, select, textarea { font-family: 'Inter', sans-serif; }
        ::selection { background: ${COLORS.gold}55; }
        @media (max-width: 480px) {
          .field-row { flex-direction: column; align-items: flex-start !important; gap: 4px !important; }
          .field-row-label { width: auto !important; }
          .field-row input, .field-row select { width: 100% !important; box-sizing: border-box; }
          .list-pills-row { padding-left: 0 !important; }
          .autofill-hint { padding-left: 0 !important; }
        }
      `}</style>

      <div style={{ maxWidth: 760, margin: "0 auto" }}>
        <Header meaningDisplay={meaningDisplay} />
        {userId && <LookupQuotaBadge count={lookupCount} limit={lookupLimit} tier={tier} isAdmin={isAdmin} meaningDisplay={meaningDisplay} />}
        <MeaningDisplayToggle value={meaningDisplay} onChange={updateMeaningDisplay} />
        <Tabs tab={tab} setTab={setTab} isAdmin={isAdmin} meaningDisplay={meaningDisplay} />

        {!loaded ? (
          <div style={{ textAlign: "center", padding: 60, color: COLORS.inkSoft }}>Đang tải…</div>
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
          />
        ) : tab === "flashcards" ? (
          <FlashcardsTab
            userId={userId}
            characterList={characterList}
            wordList={wordList}
            isAdmin={isAdmin}
            checkListAccess={checkListAccess}
            onRequireAuth={onRequireAuth}
            onViewPremium={() => setTab("premium")}
            meaningDisplay={meaningDisplay}
          />
        ) : tab === "writing" ? (
          <WritingPracticeTab
            characterList={characterList}
            isAdmin={isAdmin}
            checkListAccess={checkListAccess}
            onViewPremium={() => setTab("premium")}
            meaningDisplay={meaningDisplay}
          />
        ) : tab === "add" ? (
          <AddTab
            bushouList={bushouList}
            characterList={characterList}
            wordList={wordList}
            customWords={customWords}
            onAddCharacter={addCharacterRow}
            onAddBushou={addBushouRow}
            onUpdateCharacter={updateCharacterRow}
            onDeleteCharacter={deleteCharacterRow}
            onAddWord={addWordRow}
            onDeleteWord={deleteWordRow}
            userId={userId}
            onRequireAuth={onRequireAuth}
            onViewPremium={() => setTab("premium")}
            onQuotaUpdate={(count, limit) => {
              setLookupCount(count);
              if (typeof limit === "number") setLookupLimit(limit);
            }}
            meaningDisplay={meaningDisplay}
          />
        ) : tab === "radicals" ? (
          <RadicalsTab
            bushouList={bushouList}
            onAddBushou={addBushouRow}
            isAdmin={isAdmin}
            officialBushouKeys={officialBushouKeys}
            overrideBushouKeys={overrideBushouKeys}
            onPromoteBushou={promoteBushouToDefault}
            onWithdrawBushou={withdrawBushouFromDefault}
            meaningDisplay={meaningDisplay}
          />
        ) : tab === "hanzi" ? (
          <CharacterListPanel
            characterList={characterList}
            bushouList={bushouList}
            onDeleteCharacter={deleteCharacterRow}
            onDeleteCharacterFromOfficial={deleteCharacterFromOfficial}
            onUpdateCharacter={updateCharacterRow}
            onAddBushou={addBushouRow}
            isAdmin={isAdmin}
            officialCharKeys={officialCharKeys}
            overrideCharKeys={overrideCharKeys}
            onPromoteCharacter={promoteCharacterToDefault}
            onWithdrawCharacter={withdrawCharacterFromDefault}
            checkListAccess={checkListAccess}
            onViewPremium={() => setTab("premium")}
            meaningDisplay={meaningDisplay}
          />
        ) : tab === "vocab" ? (
          <WordListPanel
            wordList={wordList}
            characterList={characterList}
            findBushou={findBushou}
            onAddWord={addWordRow}
            onDeleteWord={deleteWordRow}
            onDeleteWordFromOfficial={deleteWordFromOfficial}
            isAdmin={isAdmin}
            officialWordKeys={officialWordKeys}
            overrideWordKeys={overrideWordKeys}
            onPromoteWord={promoteWordToDefault}
            onWithdrawWord={withdrawWordFromDefault}
            checkListAccess={checkListAccess}
            onViewPremium={() => setTab("premium")}
            meaningDisplay={meaningDisplay}
          />
        ) : tab === "premium" ? (
          <PremiumTab />
        ) : tab === "admin" ? (
          <AdminPanel isAdmin={isAdmin} allListNamesInUse={allListNamesInUse} />
        ) : null}
      </div>
    </div>
  );
}

/* ---------- Header ---------- */
function Header({ meaningDisplay }) {
  return (
    <div style={{ textAlign: "center", marginBottom: 22 }}>
      <div
        style={{
          fontFamily: "KaiTi, 'STKaiti', 'Kaiti SC', 'Noto Serif SC', serif",
          fontSize: 44,
          fontWeight: 700,
          letterSpacing: 4,
          color: COLORS.ink,
        }}
      >
        学部首学汉字
      </div>
      {meaningDisplay !== "en" && (
        <div
          style={{
            fontFamily: "Calibri, 'Segoe UI', sans-serif",
            fontSize: 17,
            fontWeight: 600,
            color: COLORS.sealDark,
            marginTop: 6,
            maxWidth: 480,
            marginLeft: "auto",
            marginRight: "auto",
            lineHeight: 1.4,
          }}
        >
          {UI_TEXT.header_subtitle.vi}
        </div>
      )}
      {meaningDisplay !== "vi" && (
        <div
          style={{
            fontFamily: meaningDisplay === "en" ? "Calibri, 'Segoe UI', sans-serif" : undefined,
            fontSize: meaningDisplay === "en" ? 17 : 12,
            fontWeight: meaningDisplay === "en" ? 600 : 400,
            color: meaningDisplay === "en" ? COLORS.sealDark : COLORS.inkSoft,
            marginTop: 6,
            letterSpacing: meaningDisplay === "en" ? undefined : 0.2,
            fontStyle: meaningDisplay === "en" ? "normal" : "italic",
            maxWidth: meaningDisplay === "en" ? 480 : 460,
            marginLeft: "auto",
            marginRight: "auto",
            lineHeight: meaningDisplay === "en" ? 1.4 : 1.5,
          }}
        >
          {UI_TEXT.header_subtitle.en}
        </div>
      )}
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
  const bg = isOut ? "rgba(166,67,46,0.08)" : isLow ? "rgba(89,89,0,0.08)" : "rgba(85,107,47,0.07)";
  return (
    <div style={{ display: "flex", justifyContent: "center", marginBottom: 20 }}>
      <div
        style={{
          display: "inline-flex",
          alignItems: "center",
          gap: 12,
          padding: "10px 22px",
          borderRadius: 999,
          border: `2px solid ${accentColor}`,
          background: bg,
        }}
      >
        <span style={{ fontSize: 17, fontWeight: 800, color: accentColor, letterSpacing: 0.4, textTransform: "uppercase" }}>
          {isAdmin ? "Admin" : tier}
        </span>
        <span style={{ width: 1, height: 20, background: accentColor, opacity: 0.35 }} />
        <span style={{ fontSize: 14, fontWeight: 600, color: COLORS.ink }}>
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
    <div style={{ display: "flex", justifyContent: "center", marginBottom: 16 }}>
      <div
        style={{
          display: "inline-flex",
          borderRadius: 999,
          border: `1px solid ${COLORS.grid}`,
          overflow: "hidden",
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
    </div>
  );
}


function Tabs({ tab, setTab, isAdmin, meaningDisplay }) {
  const items = [
    { id: "play", label: t("tab_play", meaningDisplay) },
    { id: "flashcards", label: t("tab_flashcards", meaningDisplay) },
    { id: "writing", label: t("tab_writing", meaningDisplay) },
    { id: "add", label: t("tab_add", meaningDisplay) },
    { id: "radicals", label: t("tab_radicals", meaningDisplay) },
    { id: "hanzi", label: t("tab_hanzi", meaningDisplay) },
    { id: "vocab", label: t("tab_vocab", meaningDisplay) },
    { id: "premium", label: t("tab_premium", meaningDisplay) },
  ];
  if (isAdmin) items.push({ id: "admin", label: t("tab_admin", meaningDisplay) });
  return (
    <div
      style={{
        display: "flex",
        justifyContent: "center",
        flexWrap: "wrap",
        gap: 6,
        marginBottom: 24,
        borderBottom: `1px solid ${COLORS.grid}`,
        paddingBottom: 2,
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
            color: tab === it.id ? COLORS.seal : COLORS.inkSoft,
            fontWeight: tab === it.id ? 700 : 500,
            fontSize: 14.5,
            padding: "8px 14px",
            cursor: "pointer",
          }}
        >
          {it.label}
        </button>
      ))}
    </div>
  );
}

/* ================= PLAY TAB ================= */
const REVIEW_LIST_VALUE = "__needs_review__";

function PlayTab({ characterList, wordList, bushouList, findBushou, needsReview, onMarkNeedsReview, onClearNeedsReview, isAdmin, checkListAccess, onViewPremium, meaningDisplay }) {
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
        style={{ ...selectStyle, width: 260, textAlign: "center", display: "inline-block" }}
      >
        <option value="Tất cả" style={{ background: COLORS.chipBg, color: COLORS.ink, fontWeight: 700 }}>{t("play_all_lists", meaningDisplay)}</option>
        <option value={REVIEW_LIST_VALUE} style={{ background: COLORS.chipBg, color: COLORS.ink, fontWeight: 700 }}>{t("play_review_list", meaningDisplay, needsReview.length)}</option>
        {allLists.map((l) => (
          <option key={l} value={l} style={{ background: COLORS.chipBg, color: COLORS.ink, fontWeight: 700 }}>
            {!isAdmin && checkListAccess && !checkListAccess(l) ? `🔒 ${l}` : l}
          </option>
        ))}
      </select>
      {lockedListName && (
        <ListLockedModal listName={lockedListName} onClose={() => setLockedListName(null)} onViewPremium={onViewPremium} />
      )}

      <div style={{ display: "flex", justifyContent: "center", alignItems: "center", gap: 8, marginTop: 14, flexWrap: "wrap" }}>
        <span style={{ fontSize: 14, fontWeight: 700, color: COLORS.ink }}>{t("play_difficulty_label", meaningDisplay)}</span>
        {DIFFICULTY_LEVELS.map((lvl) => (
          <button
            key={lvl.id}
            type="button"
            onClick={() => setDifficulty(lvl.id)}
            style={{
              fontSize: 14,
              fontWeight: 800,
              padding: "8px 18px",
              borderRadius: 999,
              border: `2px solid ${difficulty === lvl.id ? COLORS.seal : COLORS.grid}`,
              background: difficulty === lvl.id ? COLORS.seal : "transparent",
              color: difficulty === lvl.id ? "#FBF9EF" : COLORS.inkSoft,
              cursor: "pointer",
            }}
          >
            {lvl.label}
          </button>
        ))}
      </div>
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
        <span>Điểm: <strong style={{ color: COLORS.ink }}>{score}</strong></span>
        <span>Chuỗi đúng: <strong style={{ color: COLORS.ink }}>{streak}</strong></span>
        <span>
          {playable.length} {isWord || playable.some((p) => p.charGroups.length > 1) ? "mục" : "chữ"} có thể học
          {selectedList === REVIEW_LIST_VALUE ? " (🔁 Cần ôn lại)" : selectedList !== "Tất cả" ? ` (${selectedList})` : ""}
        </span>
      </div>

      {/* Hint card */}
      <div
        style={{
          background: COLORS.card,
          border: `1px solid ${COLORS.grid}`,
          borderRadius: 10,
          padding: "16px 20px",
          marginBottom: 22,
          textAlign: "center",
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
                <div className="pop" style={{ fontFamily: "KaiTi, 'STKaiti', 'Kaiti SC', 'Noto Serif SC', serif", fontSize: isWord ? 54 : 78, color: COLORS.bamboo }}>
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
                      fontFamily: "KaiTi, 'STKaiti', 'Kaiti SC', 'Noto Serif SC', serif",
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
          <button onClick={handleReset} className="ghost-btn" style={ghostBtnStyle}>
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
  border: `1px solid ${COLORS.grid}`,
  color: COLORS.inkSoft,
  borderRadius: 7,
  padding: "9px 16px",
  fontSize: 13.5,
  cursor: "pointer",
};

const sealBtnStyle = {
  background: COLORS.seal,
  border: "none",
  color: "#FBF9EF",
  borderRadius: 7,
  padding: "10px 22px",
  fontSize: 14,
  fontWeight: 600,
  cursor: "pointer",
  letterSpacing: 0.3,
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
function FlashcardsTab({ userId, characterList, wordList, isAdmin, checkListAccess, onRequireAuth, onViewPremium, meaningDisplay }) {
  const [selectedList, setSelectedList] = useState("Tất cả");
  const [lockedListName, setLockedListName] = useState(null);
  const [progressMap, setProgressMap] = useState(null); // null = loading
  const [queue, setQueue] = useState([]);
  const [current, setCurrent] = useState(null);
  const [flipped, setFlipped] = useState(false);
  const [sessionActive, setSessionActive] = useState(false);
  const [sessionStats, setSessionStats] = useState({ reviewed: 0, again: 0 });

  const allLists = useMemo(() => {
    const set = new Set();
    characterList.forEach((c) => getLists(c).forEach((l) => set.add(l.trim())));
    wordList.forEach((w) => (w.lists || []).forEach((l) => set.add(l.trim())));
    return Array.from(set).sort((a, b) => a.localeCompare(b, "vi"));
  }, [characterList, wordList]);

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
  }, [progressMap, characterList, wordList, selectedList]);

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
    return <div style={{ textAlign: "center", color: COLORS.inkSoft, padding: 40 }}>Đang tải…</div>;
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

          <select
            value={selectedList}
            onChange={(e) => handleListChange(e.target.value)}
            style={{ ...selectStyle, width: 260, textAlign: "center", display: "inline-block", marginBottom: 16 }}
          >
            <option value="Tất cả" style={{ background: COLORS.chipBg, color: COLORS.ink, fontWeight: 700 }}>{t("fc_all_lists", meaningDisplay)}</option>
            {allLists.map((l) => (
              <option key={l} value={l} style={{ background: COLORS.chipBg, color: COLORS.ink, fontWeight: 700 }}>
                {!isAdmin && checkListAccess && !checkListAccess(l) ? `🔒 ${l}` : l}
              </option>
            ))}
          </select>

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
              border: `2px solid ${COLORS.grid}`,
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
              <div style={{ fontFamily: "KaiTi, 'STKaiti', 'Kaiti SC', 'Noto Serif SC', serif", fontSize: current.type === "word" ? 44 : 64, color: COLORS.ink }}>
                {current.key}
              </div>
            ) : (
              <div style={{ width: "100%" }}>
                <div style={{ fontFamily: "KaiTi, 'STKaiti', 'Kaiti SC', 'Noto Serif SC', serif", fontSize: 34, color: COLORS.ink, marginBottom: 10 }}>
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
  borderRadius: 8,
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
function WritingPracticeTab({ characterList, isAdmin, checkListAccess, onViewPremium, meaningDisplay }) {
  const [selectedList, setSelectedList] = useState("Tất cả");
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
  const BRUSH_WIDTHS = { thin: 18, normal: 30, thick: 48 };
  // HanziWriter's drawingWidth is specified in its own internal SVG
  // coordinate space and gets scaled down to fit the display size, so the
  // same raw number looks much thinner there than it would as a plain
  // canvas lineWidth (which is 1 unit = 1 actual pixel, no scaling). This
  // is the recall canvas's own scale, tuned to visually match the guided
  // tracing pen at the same "Cỡ bút" selection.
  const RECALL_BRUSH_WIDTHS = { thin: 4.5, normal: 7.5, thick: 12 };
  const PREVIEW_PER_PAGE = 24; // ~8 columns x 3 rows at this layout's width

  const allLists = useMemo(() => {
    const set = new Set();
    characterList.forEach((c) => getLists(c).forEach((l) => set.add(l.trim())));
    return Array.from(set).sort((a, b) => a.localeCompare(b, "vi"));
  }, [characterList]);

  const previewChars = useMemo(() => {
    return characterList.filter((c) => selectedList === "Tất cả" || getLists(c).some((l) => l.trim() === selectedList));
  }, [characterList, selectedList]);

  const previewTotalPages = Math.max(1, Math.ceil(previewChars.length / PREVIEW_PER_PAGE));
  const previewPageItems = previewChars.slice(previewPage * PREVIEW_PER_PAGE, (previewPage + 1) * PREVIEW_PER_PAGE);

  const searchResults = useMemo(() => {
    const q = searchQuery.trim().toLowerCase();
    if (!q) return [];
    return characterList
      .filter((c) => c.char.includes(q) || c.pinyin.toLowerCase().includes(q) || c.meaning.toLowerCase().includes(q))
      .slice(0, 8);
  }, [searchQuery, characterList]);

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
    const items = characterList.filter((c) => next === "Tất cả" || getLists(c).some((l) => l.trim() === next));
    if (items.length === 0) return;
    setStatus("loading");
    setRevealOn(false);
    setQueue(shuffle(items));
    setCurrentIndex(0);
  }

  function startSession() {
    const items = characterList.filter(
      (c) => selectedList === "Tất cả" || getLists(c).some((l) => l.trim() === selectedList)
    );
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
                      fontFamily: "KaiTi, 'STKaiti', 'Kaiti SC', 'Noto Serif SC', serif",
                      fontSize: 22,
                      width: 44,
                      height: 44,
                      border: `1px solid ${COLORS.grid}`,
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

          <select
            value={selectedList}
            onChange={(e) => handleListChange(e.target.value)}
            style={{ ...selectStyle, width: 260, textAlign: "center", display: "inline-block", marginBottom: 16 }}
          >
            <option value="Tất cả" style={{ background: COLORS.chipBg, color: COLORS.ink, fontWeight: 700 }}>{t("wp_all_lists", meaningDisplay)}</option>
            {allLists.map((l) => (
              <option key={l} value={l} style={{ background: COLORS.chipBg, color: COLORS.ink, fontWeight: 700 }}>
                {!isAdmin && checkListAccess && !checkListAccess(l) ? `🔒 ${l}` : l}
              </option>
            ))}
          </select>

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
                      fontFamily: "KaiTi, 'STKaiti', 'Kaiti SC', 'Noto Serif SC', serif",
                      fontSize: 20,
                      width: 38,
                      height: 38,
                      border: `1px solid ${COLORS.grid}`,
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
                    ← Trang trước
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
                    Trang sau →
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
              <option value="Tất cả" style={{ background: COLORS.chipBg, color: COLORS.ink, fontWeight: 700 }}>Tất cả danh sách</option>
              {allLists.map((l) => (
                <option key={l} value={l} style={{ background: COLORS.chipBg, color: COLORS.ink, fontWeight: 700 }}>
                  {!isAdmin && checkListAccess && !checkListAccess(l) ? `🔒 ${l}` : l}
                </option>
              ))}
            </select>
          </div>

          <div style={{ display: "flex", gap: 6, justifyContent: "center", flexWrap: "wrap", marginBottom: 14 }}>
            {meaningDisplay !== "vi" && (
              <div style={{ padding: "3px 7px", borderRadius: 6, background: COLORS.chipBg, border: `1px solid ${COLORS.grid}`, textAlign: "center" }}>
                <div style={{ fontSize: 9, color: COLORS.inkSoft, fontWeight: 700, textTransform: "uppercase", letterSpacing: 0.3 }}>EN</div>
                <div style={{ fontSize: 11.5, color: COLORS.ink }}>{current.meaning || "—"}</div>
              </div>
            )}
            {meaningDisplay !== "en" && (
              <div style={{ padding: "3px 7px", borderRadius: 6, background: COLORS.chipBg, border: `1px solid ${COLORS.grid}`, textAlign: "center" }}>
                <div style={{ fontSize: 9, color: COLORS.inkSoft, fontWeight: 700, textTransform: "uppercase", letterSpacing: 0.3 }}>VI</div>
                <div style={{ fontSize: 11.5, color: current.meaning_vi ? COLORS.ink : COLORS.inkSoft, fontStyle: current.meaning_vi ? "normal" : "italic" }}>
                  {current.meaning_vi || "(chưa dịch)"}
                </div>
              </div>
            )}
            <div style={{ padding: "3px 7px", borderRadius: 6, background: COLORS.chipBg, border: `1px solid ${COLORS.grid}`, textAlign: "center" }}>
              <div style={{ fontSize: 9, color: COLORS.inkSoft, fontWeight: 700, textTransform: "uppercase", letterSpacing: 0.3 }}>Pinyin</div>
              <div style={{ fontSize: 11.5, color: COLORS.sealDark }}>{current.pinyin}</div>
            </div>
            {meaningDisplay !== "en" && (
              <div style={{ padding: "3px 7px", borderRadius: 6, background: COLORS.chipBg, border: `1px solid ${COLORS.grid}`, textAlign: "center" }}>
                <div style={{ fontSize: 9, color: COLORS.inkSoft, fontWeight: 700, textTransform: "uppercase", letterSpacing: 0.3 }}>Hán Việt</div>
                <div style={{ fontSize: 11.5, color: COLORS.bamboo }}>{current.sv}</div>
              </div>
            )}
          </div>

          <div style={{ display: "flex", justifyContent: "center", marginBottom: 16 }}>
            {status === "recall" ? (
              <div style={{ width: gridSize, height: gridSize, position: "relative", border: `2px solid ${COLORS.grid}`, borderRadius: 10 }}>
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
                        fontFamily: "KaiTi, 'STKaiti', 'Kaiti SC', 'Noto Serif SC', serif",
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
              <div style={{ width: gridSize, height: gridSize, position: "relative", border: `2px solid ${COLORS.grid}`, borderRadius: 10 }}>
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
                      background: "#2E8B57",
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
                    Đang tải…
                  </div>
                )}
              </div>
            ) : (
              <div style={{ width: gridSize, height: gridSize, position: "relative", border: `2px solid ${COLORS.grid}`, borderRadius: 10 }}>
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
                    <span style={{ color: "#2E8B57", fontWeight: 700 }}>{t("wp_dots_green", meaningDisplay)}</span> {t("wp_dots_start_label", meaningDisplay)}{" "}
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
                Hiện chữ đúng
              </button>
              <button
                type="button"
                onClick={undoLastDotStroke}
                disabled={completedDotStrokes.length === 0}
                className="ghost-btn"
                style={{ ...ghostBtnStyle, padding: "8px 14px", fontSize: 12.5, opacity: completedDotStrokes.length === 0 ? 0.4 : 1 }}
              >
                ↩ Xóa nét trước
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
                  background: brushSize === opt.id ? "rgba(85,107,47,0.08)" : "transparent",
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
      if (parsed.meaning && (overwrite || !meaning.trim())) setMeaning(parsed.meaning);
      if (parsed.meaning_vi && (overwrite || !meaningVi.trim())) setMeaningVi(parsed.meaning_vi);
      if (parsed.sino_vietnamese && (overwrite || !sv.trim())) setSv(parsed.sino_vietnamese);

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
      if (!parsed.pinyin && !parsed.meaning && !parsed.sino_vietnamese) {
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
      if (!charInput.trim() || !meaning.trim() || !pinyin.trim() || !sv.trim()) {
        setMessage({ type: "error", text: t("add_fill_required", meaningDisplay) });
        return;
      }
      if (characterList.some((c) => c.char === charInput.trim())) {
        setMessage({ type: "error", text: t("add_char_exists", meaningDisplay, charInput.trim()) });
        return;
      }
      const listsToSave = selectedLists.length > 0 ? selectedLists : ["Chưa phân loại"];
      const trimmedChar = charInput.trim();
      onAddCharacter({
        char: trimmedChar,
        meaning: meaning.trim(),
        meaning_vi: meaningVi.trim(),
        pinyin: pinyin.trim(),
        sv: sv.trim(),
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
              border: `1px solid ${COLORS.grid}`,
              borderRadius: 10,
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
            style={{ ...inputStyle, fontFamily: "KaiTi, 'STKaiti', 'Kaiti SC', 'Noto Serif SC', serif", fontSize: 22, width: 90, textAlign: "center" }}
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
                  <Chip info={bushouList.find((b) => b.char === ch) || { char: ch }} />
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
              style={{ ...inputStyle, fontFamily: "KaiTi, 'STKaiti', 'Kaiti SC', 'Noto Serif SC', serif", width: 140 }}
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
            <div style={{ marginTop: 12, padding: 12, background: "rgba(169,130,47,0.08)", borderRadius: 8, border: `1px dashed ${COLORS.gold}` }}>
              <div style={{ fontSize: 12.5, marginBottom: 8 }}>
                {t("add_new_component_before", meaningDisplay)}{" "}
                <strong style={{ fontFamily: "KaiTi, 'STKaiti', 'Kaiti SC', 'Noto Serif SC', serif", fontSize: 18 }}>{newCompDraft.char}</strong>{" "}
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
          <input
            value={meaning}
            onChange={(e) => setMeaning(e.target.value)}
            onKeyDown={(e) => e.key === "Enter" && handleSubmit(e)}
            placeholder="good, well"
            style={inputStyle}
          />
        </FieldRow>

        <FieldRow label={t("add_meaning_vi_label", meaningDisplay)}>
          <input
            value={meaningVi}
            onChange={(e) => setMeaningVi(e.target.value)}
            onKeyDown={(e) => e.key === "Enter" && handleSubmit(e)}
            placeholder="tốt"
            style={inputStyle}
          />
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
          <input
            value={sv}
            onChange={(e) => setSv(e.target.value)}
            onKeyDown={(e) => e.key === "Enter" && handleSubmit(e)}
            placeholder="hảo"
            style={inputStyle}
          />
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
                      background: "rgba(174,58,42,0.08)",
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
                  background: selectedLists.includes(l) ? "rgba(174,58,42,0.08)" : "transparent",
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

      <RenameListPanel
        characterList={characterList}
        wordList={wordList}
        onUpdateCharacter={onUpdateCharacter}
        onAddWord={onAddWord}
        meaningDisplay={meaningDisplay}
      />
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
    if (!parsed.pinyin && !parsed.meaning && !parsed.sino_vietnamese) throw new Error("no data returned");

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
      meaning: parsed.meaning || "",
      meaning_vi: parsed.meaning_vi || "",
      sv: parsed.sino_vietnamese || "",
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
            await onAddWord({
              word: item,
              chars: Array.from(item),
              pinyin: wordParsed.pinyin || "",
              meaning: wordParsed.meaning || "",
              meaning_vi: wordParsed.meaning_vi || "",
              sv: wordParsed.sino_vietnamese || "",
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
        background: "rgba(89,89,0,0.05)",
        border: `1px dashed ${COLORS.gold}`,
        borderRadius: 8,
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
            style={{ ...inputStyle, width: "100%", boxSizing: "border-box", fontFamily: "KaiTi, 'STKaiti', 'Kaiti SC', 'Noto Serif SC', serif", fontSize: 15, resize: "vertical", marginBottom: 6, whiteSpace: "pre" }}
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
                      background: "rgba(85,107,47,0.08)",
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
                      background: selectedLists.includes(l) ? "rgba(85,107,47,0.08)" : "transparent",
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
function RenameListPanel({ characterList, wordList, onUpdateCharacter, onAddWord, meaningDisplay }) {
  const [expanded, setExpanded] = useState(false);
  const [oldName, setOldName] = useState("");
  const [newName, setNewName] = useState("");
  const [status, setStatus] = useState("idle"); // idle | running | done
  const [progress, setProgress] = useState({ done: 0, total: 0 });
  const [message, setMessage] = useState(null);

  const allLists = useMemo(() => {
    const set = new Set();
    characterList.forEach((c) => getLists(c).forEach((l) => set.add(l.trim())));
    wordList.forEach((w) => (w.lists || []).forEach((l) => set.add(l.trim())));
    return Array.from(set).sort((a, b) => a.localeCompare(b, "vi"));
  }, [characterList, wordList]);

  async function handleRename() {
    setMessage(null);
    const trimmedOld = oldName.trim();
    const trimmedNew = newName.trim();
    if (!trimmedOld) {
      setMessage({ type: "error", text: t("rename_need_old", meaningDisplay) });
      return;
    }
    if (!trimmedNew) {
      setMessage({ type: "error", text: t("rename_need_new", meaningDisplay) });
      return;
    }
    if (trimmedOld === trimmedNew) {
      setMessage({ type: "error", text: t("rename_same_name", meaningDisplay) });
      return;
    }

    const affectedChars = characterList.filter((c) => getLists(c).includes(trimmedOld));
    const affectedWords = wordList.filter((w) => (w.lists || []).includes(trimmedOld));
    const total = affectedChars.length + affectedWords.length;

    if (total === 0) {
      setMessage({ type: "error", text: t("rename_none_found", meaningDisplay, trimmedOld) });
      return;
    }

    setStatus("running");
    setProgress({ done: 0, total });

    let done = 0;
    for (const c of affectedChars) {
      const lists = getLists(c).map((l) => (l === trimmedOld ? trimmedNew : l));
      await onUpdateCharacter(c.char, { lists });
      done += 1;
      setProgress({ done, total });
    }
    for (const w of affectedWords) {
      const lists = (w.lists || []).map((l) => (l === trimmedOld ? trimmedNew : l));
      await onAddWord({ ...w, lists });
      done += 1;
      setProgress({ done, total });
    }

    setStatus("done");
    setMessage({ type: "success", text: t("rename_success", meaningDisplay, trimmedOld, trimmedNew, total) });
    setOldName("");
    setNewName("");
  }

  return (
    <div
      style={{
        background: "rgba(89,89,0,0.05)",
        border: `1px dashed ${COLORS.gold}`,
        borderRadius: 8,
        padding: "12px 14px",
        marginBottom: 18,
      }}
    >
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
        {expanded ? "▲" : "▼"} {t("rename_toggle", meaningDisplay)}
      </button>

      {expanded && (
        <div style={{ marginTop: 12 }}>
          <div style={{ fontSize: 11.5, color: COLORS.inkSoft, marginBottom: 10 }}>
            {t("rename_description", meaningDisplay)}
          </div>

          <div style={{ display: "flex", gap: 8, marginBottom: 10, flexWrap: "wrap" }}>
            <select
              value={oldName}
              onChange={(e) => setOldName(e.target.value)}
              disabled={status === "running"}
              style={{ ...selectStyle, width: 180 }}
            >
              <option value="" style={{ background: COLORS.chipBg, color: COLORS.ink, fontWeight: 700 }}>{t("rename_choose_list", meaningDisplay)}</option>
              {allLists.map((l) => (
                <option key={l} value={l} style={{ background: COLORS.chipBg, color: COLORS.ink, fontWeight: 700 }}>
                  {l}
                </option>
              ))}
            </select>
            <span style={{ alignSelf: "center", color: COLORS.inkSoft }}>→</span>
            <input
              value={newName}
              onChange={(e) => setNewName(e.target.value)}
              disabled={status === "running"}
              placeholder={t("rename_new_name_placeholder", meaningDisplay)}
              style={{ ...inputStyle, maxWidth: 180 }}
            />
            <button
              type="button"
              onClick={handleRename}
              disabled={status === "running"}
              className="seal-btn"
              style={{ ...sealBtnStyle, padding: "8px 16px", fontSize: 13 }}
            >
              {status === "running" ? t("rename_renaming", meaningDisplay) : t("rename_button", meaningDisplay)}
            </button>
          </div>

          {status === "running" && progress.total > 0 && (
            <div style={{ fontSize: 12, color: COLORS.inkSoft, marginBottom: 8 }}>
              {progress.done} / {progress.total}
            </div>
          )}

          {message && (
            <div style={{ fontSize: 12.5, fontWeight: 600, color: message.type === "error" ? COLORS.error : COLORS.bamboo }}>
              {message.text}
            </div>
          )}
        </div>
      )}
    </div>
  );
}

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
      if (parsed.meaning) setMeaning(parsed.meaning);
      if (parsed.meaning_vi) setMeaningVi(parsed.meaning_vi);
      if (parsed.sino_vietnamese) setSv(parsed.sino_vietnamese);
      if (!parsed.pinyin && !parsed.meaning) {
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
    if (!pinyin.trim() || !meaning.trim()) {
      setMessage({ type: "error", text: t("word_need_pinyin_meaning", meaningDisplay) });
      return;
    }
    if (wordList.some((w) => w.word === word)) {
      setMessage({ type: "error", text: t("word_exists", meaningDisplay, word) });
      return;
    }
    const listsToSave = selectedLists.length > 0 ? selectedLists : ["Chưa phân loại"];
    onAddWord({ word, chars, pinyin: pinyin.trim(), meaning: meaning.trim(), meaning_vi: meaningVi.trim(), sv: sv.trim(), lists: listsToSave });
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
        background: "rgba(89,89,0,0.05)",
        border: `1px dashed ${COLORS.gold}`,
        borderRadius: 8,
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
              style={{ ...inputStyle, fontFamily: "KaiTi, 'STKaiti', 'Kaiti SC', 'Noto Serif SC', serif", fontSize: 20 }}
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
                    <span style={{ fontFamily: "KaiTi, 'STKaiti', 'Kaiti SC', 'Noto Serif SC', serif", fontSize: 20 }}>{ch}</span>
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
            <input value={meaning} onChange={(e) => setMeaning(e.target.value)} placeholder={t("word_auto_or_manual", meaningDisplay)} style={inputStyle} />
          </FieldRow>
          <FieldRow label={t("word_meaning_vi_label", meaningDisplay)}>
            <input value={meaningVi} onChange={(e) => setMeaningVi(e.target.value)} placeholder={t("word_auto_or_manual", meaningDisplay)} style={inputStyle} />
          </FieldRow>
          <FieldRow label={t("word_hanviet_label", meaningDisplay)}>
            <input value={sv} onChange={(e) => setSv(e.target.value)} placeholder={t("word_auto_or_manual", meaningDisplay)} style={inputStyle} />
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
                        background: "rgba(85,107,47,0.08)",
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
                        background: selectedLists.includes(l) ? "rgba(85,107,47,0.08)" : "transparent",
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
const TIER_PRESETS = { Free: 100, Silver: 500, Titan: 2000, Gold: 5000, Platinum: 15000 };
// "Enrolled Course" isn't a fixed-limit preset like the others -- its
// limit is set manually per course/student, and it carries an extra
// course_name field the others don't use.
const ALL_TIERS = [...Object.keys(TIER_PRESETS), "Enrolled Course"];

/* ---------- Sales/info page about paid tiers and courses. Visible to
   everyone, including guests -- this is the conversion page the
   quota-exhausted and locked-list popups link to. Static content for now;
   revise the copy freely, it's just plain text/JSX below. ---------- */
function PremiumTab() {
  const tierCard = (name, limit, blurb) => (
    <div
      style={{
        background: COLORS.card,
        border: `1px solid ${COLORS.grid}`,
        borderRadius: 10,
        padding: "16px 18px",
        marginBottom: 10,
      }}
    >
      <div style={{ display: "flex", alignItems: "baseline", gap: 10, marginBottom: 4 }}>
        <span style={{ fontSize: 16, fontWeight: 800, color: COLORS.seal, textTransform: "uppercase" }}>{name}</span>
        <span style={{ fontSize: 12.5, color: COLORS.inkSoft }}>{limit} lượt tra cứu tự động</span>
      </div>
      <div style={{ fontSize: 12.5, color: COLORS.inkSoft, lineHeight: 1.5 }}>{blurb}</div>
    </div>
  );

  return (
    <div style={{ maxWidth: 560, margin: "0 auto" }}>
      <div style={{ fontSize: 12.5, fontWeight: 600, color: COLORS.gold, marginBottom: 4, textTransform: "uppercase", letterSpacing: 0.8, textAlign: "center" }}>
        Thành viên trả phí & Khóa học
      </div>
      <div style={{ fontSize: 13, color: COLORS.inkSoft, textAlign: "center", marginBottom: 24, lineHeight: 1.6 }}>
        Học Chữ Hán cung cấp miễn phí bộ dữ liệu cơ bản để bạn bắt đầu học chữ Hán mọi lúc. Với gói thành viên trả
        phí, bạn sẽ được tra cứu tự động nhiều hơn và mở khóa các danh sách từ vựng chuyên sâu do chúng tôi biên
        soạn riêng.
      </div>

      <div style={{ fontSize: 12, fontWeight: 700, color: COLORS.ink, textTransform: "uppercase", letterSpacing: 0.5, marginBottom: 10 }}>
        Bảng gói thành viên
      </div>
      {tierCard("Free", "100", "Truy cập toàn bộ dữ liệu cơ bản.")}
      {tierCard("Silver", "500", "Bao gồm các danh sách từ vựng nâng cao dành riêng cho gói Silver trở lên.")}
      {tierCard("Titan", "2.000", "Bao gồm các danh sách từ vựng nâng cao dành riêng cho gói Titan trở lên.")}
      {tierCard("Gold", "5.000", "Bao gồm các danh sách từ vựng nâng cao dành riêng cho gói Gold trở lên.")}
      {tierCard("Platinum", "15.000", "Bao gồm toàn bộ danh sách từ vựng nâng cao hiện có.")}

      <div style={{ fontSize: 12, fontWeight: 700, color: COLORS.ink, textTransform: "uppercase", letterSpacing: 0.5, marginTop: 20, marginBottom: 10 }}>
        Khóa học riêng
      </div>
      <div style={{ fontSize: 12.5, color: COLORS.inkSoft, lineHeight: 1.6, marginBottom: 20 }}>
        Ngoài các gói thành viên, chúng tôi cũng tổ chức các khóa học với bộ từ vựng được biên soạn riêng cho từng
        khóa. Nếu bạn đang theo học một khóa cụ thể, tài khoản của bạn sẽ được cấp quyền truy cập vào danh sách từ
        vựng riêng của khóa đó.
      </div>

      <div style={{ fontSize: 12, fontWeight: 700, color: COLORS.ink, textTransform: "uppercase", letterSpacing: 0.5, marginBottom: 10 }}>
        Cách nâng cấp
      </div>
      <div style={{ fontSize: 12.5, color: COLORS.inkSoft, lineHeight: 1.7 }}>
        Hiện tại, việc nâng cấp được thực hiện thủ công:
        <ol style={{ margin: "8px 0", paddingLeft: 20 }}>
          <li>Chuyển khoản theo thông tin: [điền thông tin chuyển khoản]</li>
          <li>Ghi chú nội dung chuyển khoản: [email tài khoản của bạn]</li>
          <li>Tài khoản của bạn sẽ được nâng cấp trong vòng [điền thời gian]</li>
        </ol>
        Mọi thắc mắc xin liên hệ: [điền email hoặc kênh liên hệ]
      </div>
    </div>
  );
}

function AdminPanel({ isAdmin, allListNamesInUse }) {
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

  useEffect(() => {
    if (!isAdmin) return;
    loadUsers();
    loadListSettings();
    // eslint-disable-next-line react-hooks/exhaustive-deps
  }, [isAdmin]);

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
      setListMessage({ type: "error", text: "Không thể lưu: " + error.message });
      return;
    }
    setListSettings((prev) => {
      const without = prev.filter((s) => s.name !== name);
      return [...without, { name, admin_only: editAdminOnly, allowed_tiers: editAllowedTiers }];
    });
    setEditingListName(null);
    setListMessage({ type: "success", text: "Đã lưu." });
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
    setEditLimit(String(u.lookup_limit != null ? u.lookup_limit : 100));
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
      setMessage({ type: "error", text: "Giới hạn không hợp lệ." });
      return;
    }
    const isCourse = editTier === "Enrolled Course";
    if (isCourse && !editCourseName.trim()) {
      setMessage({ type: "error", text: "Vui lòng nhập tên khóa học." });
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
    if (!window.confirm("Đặt lại lượt tra cứu về 0 cho người dùng này?")) return;
    const { error } = await supabase.from("profiles").update({ lookup_count: 0 }).eq("user_id", userId);
    if (error) {
      setMessage({ type: "error", text: "Không thể đặt lại: " + error.message });
      return;
    }
    setUsers((prev) => prev.map((u) => (u.user_id === userId ? { ...u, lookup_count: 0 } : u)));
    setMessage({ type: "success", text: "Đã đặt lại." });
    setTimeout(() => setMessage(null), 2500);
  }

  async function toggleDisabled(u) {
    const next = !u.disabled;
    if (next && !window.confirm(`Vô hiệu hóa tài khoản "${u.email || u.user_id}"? Người này sẽ không thể tra cứu tự động cho đến khi được kích hoạt lại.`)) return;
    const { error } = await supabase.from("profiles").update({ disabled: next }).eq("user_id", u.user_id);
    if (error) {
      setMessage({ type: "error", text: "Không thể cập nhật: " + error.message });
      return;
    }
    setUsers((prev) => prev.map((x) => (x.user_id === u.user_id ? { ...x, disabled: next } : x)));
    setMessage({ type: "success", text: next ? "Đã vô hiệu hóa." : "Đã kích hoạt lại." });
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

  return (
    <div>
      <div style={{ fontSize: 12.5, fontWeight: 600, color: COLORS.gold, marginBottom: 12, textTransform: "uppercase", letterSpacing: 0.8, textAlign: "center" }}>
        Quản trị người dùng
      </div>

      <div style={{ display: "flex", justifyContent: "center", flexWrap: "wrap", gap: 8, marginBottom: 10 }}>
        <input
          value={query}
          onChange={(e) => setQuery(e.target.value)}
          placeholder="Tìm theo email…"
          style={{ ...inputStyle, width: 220, textAlign: "center" }}
        />
        <select value={tierFilter} onChange={(e) => setTierFilter(e.target.value)} style={{ ...selectStyle, width: 150, flex: "none" }}>
          <option value="Tất cả" style={{ background: COLORS.chipBg, color: COLORS.ink, fontWeight: 700 }}>Tất cả gói</option>
          {ALL_TIERS.map((t) => (
            <option key={t} value={t} style={{ background: COLORS.chipBg, color: COLORS.ink, fontWeight: 700 }}>
              {t}
            </option>
          ))}
        </select>
        <select value={statusFilter} onChange={(e) => setStatusFilter(e.target.value)} style={{ ...selectStyle, width: 170, flex: "none" }}>
          <option value="all" style={{ background: COLORS.chipBg, color: COLORS.ink, fontWeight: 700 }}>Tất cả trạng thái</option>
          <option value="enabled" style={{ background: COLORS.chipBg, color: COLORS.ink, fontWeight: 700 }}>Đang hoạt động</option>
          <option value="disabled" style={{ background: COLORS.chipBg, color: COLORS.ink, fontWeight: 700 }}>Đã vô hiệu hóa</option>
        </select>
        {existingCourseNames.length > 0 && (
          <select value={courseFilter} onChange={(e) => setCourseFilter(e.target.value)} style={{ ...selectStyle, width: 170, flex: "none" }}>
            <option value="Tất cả" style={{ background: COLORS.chipBg, color: COLORS.ink, fontWeight: 700 }}>Tất cả khóa học</option>
            {existingCourseNames.map((c) => (
              <option key={c} value={c} style={{ background: COLORS.chipBg, color: COLORS.ink, fontWeight: 700 }}>
                {c}
              </option>
            ))}
          </select>
        )}
        <button type="button" onClick={loadUsers} className="ghost-btn" style={{ ...ghostBtnStyle, padding: "8px 14px", fontSize: 12.5 }}>
          ⟳ Làm mới
        </button>
      </div>

      <div style={{ fontSize: 11.5, color: COLORS.inkSoft, textAlign: "center", marginBottom: 14 }}>
        {filtered.length} / {users.length} người dùng
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
        <div style={{ textAlign: "center", color: COLORS.inkSoft, padding: 30 }}>Đang tải…</div>
      ) : (
        <div style={{ display: "flex", flexDirection: "column", gap: 8 }}>
          {filtered.map((u) => (
            <div
              key={u.user_id}
              style={{
                background: COLORS.card,
                border: `1px solid ${COLORS.grid}`,
                borderRadius: 8,
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
                        placeholder="Tên khóa học…"
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
                    Lưu
                  </button>
                  <button
                    type="button"
                    onClick={() => setEditingId(null)}
                    className="ghost-btn"
                    style={{ ...ghostBtnStyle, padding: "6px 14px", fontSize: 12 }}
                  >
                    Hủy
                  </button>
                </div>
              ) : (
                <div style={{ display: "flex", flexWrap: "wrap", alignItems: "center", gap: 10 }}>
                  <div style={{ flex: "1 1 200px", fontSize: 13, color: COLORS.ink, fontWeight: 600 }}>
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
                          background: "rgba(166,67,46,0.08)",
                          color: COLORS.error,
                        }}
                      >
                        Đã vô hiệu hóa
                      </span>
                    )}
                  </div>
                  <div style={{ fontSize: 12, color: COLORS.sealDark, minWidth: 60 }}>
                    {u.tier || "Free"}
                    {u.tier === "Enrolled Course" && u.course_name && (
                      <span style={{ color: COLORS.inkSoft, fontWeight: 500 }}> · {u.course_name}</span>
                    )}
                  </div>
                  <div style={{ fontSize: 12, color: COLORS.inkSoft, minWidth: 80 }}>
                    {u.lookup_count ?? 0} / {u.lookup_limit ?? 100}
                  </div>
                  <button
                    type="button"
                    onClick={() => startEdit(u)}
                    className="ghost-btn"
                    style={{ ...ghostBtnStyle, padding: "5px 10px", fontSize: 11.5 }}
                  >
                    Sửa
                  </button>
                  <button
                    type="button"
                    onClick={() => resetUsage(u.user_id)}
                    className="ghost-btn"
                    style={{ ...ghostBtnStyle, padding: "5px 10px", fontSize: 11.5 }}
                  >
                    Đặt lại về 0
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
                      {u.disabled ? "✓ Kích hoạt lại" : "🚫 Vô hiệu hóa"}
                    </button>
                  )}
                </div>
              )}
            </div>
          ))}
          {filtered.length === 0 && (
            <div style={{ textAlign: "center", color: COLORS.inkSoft, padding: 20 }}>Không tìm thấy người dùng.</div>
          )}
        </div>
      )}

      <div style={{ marginTop: 32, paddingTop: 22, borderTop: `1px dashed ${COLORS.grid}` }}>
        <div style={{ fontSize: 12.5, fontWeight: 600, color: COLORS.gold, marginBottom: 12, textTransform: "uppercase", letterSpacing: 0.8, textAlign: "center" }}>
          Quản lý danh sách
        </div>
        <div style={{ fontSize: 11.5, color: COLORS.inkSoft, textAlign: "center", marginBottom: 16, lineHeight: 1.5 }}>
          Danh sách chưa cấu hình bên dưới mặc định mở cho mọi người. "Chỉ admin" ẩn hoàn toàn khỏi người dùng
          thường. Chọn gói và/hoặc gán khóa học cụ thể để giới hạn quyền xem nội dung (tên danh sách vẫn hiển thị
          cho mọi người, trừ khi chọn "Chỉ admin").
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
          <div style={{ textAlign: "center", color: COLORS.inkSoft, padding: 20 }}>Đang tải…</div>
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
                    border: `1px solid ${COLORS.grid}`,
                    borderRadius: 8,
                    padding: "10px 14px",
                  }}
                >
                  {isEditing ? (
                    <div>
                      <div style={{ display: "flex", flexWrap: "wrap", alignItems: "center", gap: 10, marginBottom: 10 }}>
                        <div style={{ flex: "1 1 160px", fontSize: 13, color: COLORS.ink, fontWeight: 600 }}>{name}</div>
                        <label style={{ display: "flex", alignItems: "center", gap: 6, fontSize: 12.5, color: COLORS.ink, cursor: "pointer" }}>
                          <input type="checkbox" checked={editAdminOnly} onChange={(e) => setEditAdminOnly(e.target.checked)} />
                          Chỉ admin
                        </label>
                      </div>

                      {!editAdminOnly && (
                        <>
                          <div style={{ fontSize: 11, color: COLORS.inkSoft, marginBottom: 6 }}>Gói được phép xem:</div>
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
                                  background: editAllowedTiers.includes(t) ? "rgba(85,107,47,0.08)" : "transparent",
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
                            Cấp riêng cho khóa học cụ thể (dành cho gói "Enrolled Course"):
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
                                    background: "rgba(85,107,47,0.08)",
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
                              placeholder="Tên khóa học…"
                              style={{ ...inputStyle, fontSize: 12.5, padding: "6px 10px", width: 200 }}
                            />
                            <button
                              type="button"
                              onClick={() => addCourseGrant(name)}
                              className="ghost-btn"
                              style={{ ...ghostBtnStyle, padding: "6px 12px", fontSize: 12 }}
                            >
                              + Thêm
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
                          Lưu
                        </button>
                        <button
                          type="button"
                          onClick={() => setEditingListName(null)}
                          className="ghost-btn"
                          style={{ ...ghostBtnStyle, padding: "6px 14px", fontSize: 12 }}
                        >
                          Hủy
                        </button>
                      </div>
                    </div>
                  ) : (
                    <div style={{ display: "flex", flexWrap: "wrap", alignItems: "center", gap: 10 }}>
                      <div style={{ flex: "1 1 160px", fontSize: 13, color: COLORS.ink, fontWeight: 600 }}>{name}</div>
                      <div style={{ fontSize: 12, color: COLORS.inkSoft }}>
                        {isOpen
                          ? "Mở cho tất cả"
                          : setting.admin_only
                          ? "🔒 Chỉ admin"
                          : [
                              ...(setting.allowed_tiers || []),
                              ...(courseGrants.length > 0 ? [`${courseGrants.length} khóa học`] : []),
                            ].join(", ") || "Không ai được xem"}
                      </div>
                      <button
                        type="button"
                        onClick={() => startEditList(name)}
                        className="ghost-btn"
                        style={{ ...ghostBtnStyle, padding: "5px 10px", fontSize: 11.5 }}
                      >
                        Sửa
                      </button>
                    </div>
                  )}
                </div>
              );
            })}
            {allListNamesInUse.length === 0 && (
              <div style={{ textAlign: "center", color: COLORS.inkSoft, padding: 20 }}>Chưa có danh sách nào.</div>
            )}
          </div>
        )}
      </div>
    </div>
  );
}

function WordListPanel({ wordList, characterList, findBushou, onAddWord, onDeleteWord, onDeleteWordFromOfficial, isAdmin, officialWordKeys, overrideWordKeys, onPromoteWord, onWithdrawWord, checkListAccess, onViewPremium, meaningDisplay }) {
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
        {t("vocab_panel_title", meaningDisplay)}
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
                  {!isAdmin && checkListAccess && !checkListAccess(l) ? `🔒 ${l}` : l}
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
        <div style={{ fontFamily: "KaiTi, 'STKaiti', 'Kaiti SC', 'Noto Serif SC', serif", fontSize: 18, marginBottom: 6 }}>{w.word}</div>
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
                  background: "rgba(85,107,47,0.08)",
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
                  background: selectedLists.includes(l) ? "rgba(85,107,47,0.08)" : "transparent",
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
        border: `1px solid ${COLORS.grid}`,
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
          border: `1px solid ${COLORS.grid}`,
          borderRadius: "50%",
          background: COLORS.chipBg,
          color: COLORS.sealDark,
          cursor: "pointer",
        }}
      >
        🔍
      </button>

      <div style={{ position: "absolute", top: 6, right: 6, display: "flex", gap: 4 }}>
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
      </div>

      <div
        onClick={() => setZoomed(true)}
        title="Bấm để phóng to"
        style={{ fontFamily: "KaiTi, 'STKaiti', 'Kaiti SC', 'Noto Serif SC', serif", fontSize: 20, color: COLORS.ink, marginBottom: 2, marginTop: 6, cursor: "pointer" }}
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
            background: !hasOverride && isOfficial ? "rgba(89,89,0,0.12)" : "rgba(89,89,0,0.06)",
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
                <span style={{ fontFamily: "KaiTi, 'STKaiti', 'Kaiti SC', 'Noto Serif SC', serif", color: COLORS.ink, fontSize: 13 }}>{ch}</span>
                {" = "}
                {found.components.map((comp, ci) => (
                  <span key={ci}>
                    <span
                      title={findBushou ? `${findBushou(comp).pinyin} · ${findBushou(comp).meaning} · HV: ${findBushou(comp).sv}` : undefined}
                      style={{ fontFamily: "KaiTi, 'STKaiti', 'Kaiti SC', 'Noto Serif SC', serif", color: COLORS.gold, fontSize: 13 }}
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
        <WordZoomModal w={w} characterList={characterList} findBushou={findBushou} onClose={() => setZoomed(false)} meaningDisplay={meaningDisplay} />
      )}
    </div>
  );
}

/* ---------- Full-screen study view for one word: one mizige box per
   character (same idea as the Play tab's multi-box build area), plus
   pinyin/meaning/Hán Việt and each character's own bushou breakdown. ---------- */
function WordZoomModal({ w, characterList, findBushou, onClose, meaningDisplay }) {
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
            border: `1px solid ${COLORS.grid}`,
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
              <div style={{ fontFamily: "KaiTi, 'STKaiti', 'Kaiti SC', 'Noto Serif SC', serif", fontSize: Math.round(boxSize * 0.6), color: COLORS.ink }}>
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
              Bộ thủ cấu thành
            </div>
            {Array.from(new Set(chars)).map((ch, i) => {
              const found = characterList.find((c) => c.char === ch);
              if (!found || !found.components || found.components.length === 0) return null;
              return (
                <div key={i} style={{ marginBottom: 14 }}>
                  <div style={{ fontSize: 13, color: COLORS.ink, marginBottom: 8, fontWeight: 600, display: "flex", alignItems: "center", justifyContent: "center", gap: 8 }}>
                    <span style={{ fontFamily: "KaiTi, 'STKaiti', 'Kaiti SC', 'Noto Serif SC', serif", fontSize: 18 }}>{ch}</span>
                    <span>({found.pinyin} · {found.meaning})</span>
                    <button
                      type="button"
                      onClick={() => setStrokeChar(ch)}
                      title={`Xem thứ tự nét bút của ${ch}`}
                      style={{
                        width: 20,
                        height: 20,
                        lineHeight: "18px",
                        padding: 0,
                        fontSize: 11,
                        border: `1px solid ${COLORS.grid}`,
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
                      <Chip key={ci} info={findBushou(comp)} big disabled />
                    ))}
                  </div>
                  <div style={{ display: "flex", flexDirection: "column", gap: 3, textAlign: "left" }}>
                    {found.components.map((comp, ci) => {
                      const info = findBushou(comp);
                      return (
                        <div key={ci} style={{ fontSize: 12, color: COLORS.inkSoft }}>
                          <span style={{ fontFamily: "KaiTi, 'STKaiti', 'Kaiti SC', 'Noto Serif SC', serif", fontSize: 15, color: COLORS.ink }}>{comp}</span>
                          {" — "}
                          {info.pinyin} · {info.meaning} · HV: {info.sv}
                        </div>
                      );
                    })}
                  </div>
                </div>
              );
            })}
          </div>
        )}

        {strokeChar && <StrokeOrderModal char={strokeChar} onClose={() => setStrokeChar(null)} />}
      </div>
    </div>
  );
}

/* ---------- List function: browse every character already in the database ---------- */
function CharacterListPanel({ characterList, bushouList, onDeleteCharacter, onDeleteCharacterFromOfficial, onUpdateCharacter, onAddBushou, isAdmin, officialCharKeys, overrideCharKeys, onPromoteCharacter, onWithdrawCharacter, checkListAccess, onViewPremium, meaningDisplay }) {
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
        {t("hanzi_panel_title", meaningDisplay)}
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
                {!isAdmin && checkListAccess && !checkListAccess(l) ? `🔒 ${l}` : l}
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
        border: `1px solid ${COLORS.grid}`,
        borderRadius: 8,
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
            border: `1px solid ${COLORS.grid}`,
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
              border: `1px solid ${COLORS.grid}`,
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
              border: `1px solid ${COLORS.grid}`,
              borderRadius: "50%",
              background: COLORS.chipBg,
              color: COLORS.error,
              cursor: "pointer",
            }}
          >
            ✕
          </button>
        </div>
      )}

      {mode === "edit" ? (
        <div style={{ textAlign: "left" }}>
          <div style={{ fontFamily: "KaiTi, 'STKaiti', 'Kaiti SC', 'Noto Serif SC', serif", fontSize: 26, color: COLORS.ink, textAlign: "center", marginBottom: 8 }}>
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
                    background: "rgba(174,58,42,0.08)",
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

          <label style={{ fontSize: 10, color: COLORS.inkSoft, display: "block", marginBottom: 4 }}>Bộ thủ cấu thành</label>
          {components.length > 0 && (
            <div style={{ display: "flex", flexWrap: "wrap", gap: 4, marginBottom: 6 }}>
              {components.map((ch, i) => (
                <div key={i} style={{ display: "flex", alignItems: "center", gap: 2 }}>
                  <span
                    style={{
                      fontFamily: "KaiTi, 'STKaiti', 'Kaiti SC', 'Noto Serif SC', serif",
                      fontSize: 14,
                      color: COLORS.gold,
                      padding: "1px 5px",
                      border: `1px solid ${COLORS.grid}`,
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
              style={{ ...inputStyle, fontFamily: "KaiTi, 'STKaiti', 'Kaiti SC', 'Noto Serif SC', serif", fontSize: 12.5, padding: "5px 8px" }}
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
            <div style={{ padding: 8, background: "rgba(169,130,47,0.08)", borderRadius: 6, border: `1px dashed ${COLORS.gold}`, marginBottom: 6 }}>
              <div style={{ fontSize: 11, marginBottom: 6 }}>
                Bộ thủ <strong style={{ fontFamily: "KaiTi, 'STKaiti', 'Kaiti SC', 'Noto Serif SC', serif" }}>{newCompDraft.char}</strong> chưa có — điền thông tin:
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
            style={{ fontFamily: "KaiTi, 'STKaiti', 'Kaiti SC', 'Noto Serif SC', serif", fontSize: 32, color: COLORS.ink, cursor: "pointer", marginTop: 10 }}
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
                background: !hasOverride && isOfficial ? "rgba(89,89,0,0.12)" : "rgba(89,89,0,0.06)",
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
                    fontFamily: "KaiTi, 'STKaiti', 'Kaiti SC', 'Noto Serif SC', serif",
                    fontSize: 16,
                    color: COLORS.gold,
                    padding: "2px 5px",
                    border: `1px solid ${COLORS.grid}`,
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

      {zoomed && <CharacterZoomModal c={c} findBushou={findBushou} onClose={() => setZoomed(false)} meaningDisplay={meaningDisplay} />}
    </div>
  );
}

/* ---------- Full-screen study view for one character: big glyph in the
   same mizige grid used during Play, plus pinyin/meaning/Hán Việt and each
   component's own details, all at a much larger size than the card. ---------- */
function CharacterZoomModal({ c, findBushou, onClose, meaningDisplay }) {
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
            border: `1px solid ${COLORS.grid}`,
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
            <div style={{ fontFamily: "KaiTi, 'STKaiti', 'Kaiti SC', 'Noto Serif SC', serif", fontSize: 130, color: COLORS.ink }}>{c.char}</div>
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
            ✍️ Xem thứ tự nét bút
          </button>
        </div>

        {strokeOrderOpen && <StrokeOrderModal char={c.char} onClose={() => setStrokeOrderOpen(false)} />}

        {c.components && c.components.length > 0 && (
          <div style={{ borderTop: `1px dashed ${COLORS.grid}`, paddingTop: 16 }}>
            <div style={{ fontSize: 11, color: COLORS.inkSoft, textTransform: "uppercase", letterSpacing: 0.5, fontWeight: 700, marginBottom: 10 }}>
              Bộ thủ cấu thành
            </div>
            <div style={{ display: "flex", justifyContent: "center", flexWrap: "wrap", gap: 10, marginBottom: 12 }}>
              {c.components.map((comp, i) => (
                <Chip key={i} info={findBushou(comp)} big disabled />
              ))}
            </div>
            <div style={{ display: "flex", flexDirection: "column", gap: 4, textAlign: "left" }}>
              {c.components.map((comp, i) => {
                const info = findBushou(comp);
                return (
                  <div key={i} style={{ fontSize: 12.5, color: COLORS.inkSoft }}>
                    <span style={{ fontFamily: "KaiTi, 'STKaiti', 'Kaiti SC', 'Noto Serif SC', serif", fontSize: 16, color: COLORS.ink }}>{comp}</span>
                    {" — "}
                    {info.pinyin} · {info.meaning} · HV: {info.sv}
                  </div>
                );
              })}
            </div>
          </div>
        )}
      </div>
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
  border: `1px solid ${COLORS.grid}`,
  borderRadius: 10,
  padding: "22px 24px",
};

const inputStyle = {
  border: `1px solid ${COLORS.grid}`,
  borderRadius: 6,
  padding: "8px 10px",
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
  fontWeight: 700,
  border: `1.5px solid ${COLORS.seal}`,
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
function RadicalsTab({ bushouList, onAddBushou, isAdmin, officialBushouKeys, overrideBushouKeys, onPromoteBushou, onWithdrawBushou, meaningDisplay }) {
  const [query, setQuery] = useState("");
  const filtered = bushouList.filter((b) => {
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
          border: `1px solid ${COLORS.grid}`,
          borderRadius: 10,
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

      <div style={{ textAlign: "center", marginBottom: 16 }}>
        <input
          value={query}
          onChange={(e) => setQuery(e.target.value)}
          placeholder={t("radicals_search_placeholder", meaningDisplay)}
          style={{ ...inputStyle, width: 340, textAlign: "center" }}
        />
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
              {strokeCount === "?" ? "chưa xác định số nét" : `${strokeCount} nét`}
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
function RadicalCard({ b, onAddBushou, isAdmin, isOfficial, hasOverride, onPromoteBushou, onWithdrawBushou, meaningDisplay }) {
  const [mode, setMode] = useState("view"); // view | edit
  const [strokeOrderOpen, setStrokeOrderOpen] = useState(false);
  const [pinyin, setPinyin] = useState(b.pinyin);
  const [meaning, setMeaning] = useState(b.meaning);
  const [sv, setSv] = useState(b.sv);
  const [strokes, setStrokes] = useState(typeof b.strokes === "number" ? String(b.strokes) : "");
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
    setMode("edit");
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
      });
    setMode("view");
  }

  return (
    <div
      style={{
        background: COLORS.card,
        border: `1px solid ${COLORS.grid}`,
        borderRadius: 8,
        padding: "12px 10px",
        textAlign: "center",
        position: "relative",
      }}
    >
      {mode === "view" && (
        <button
          type="button"
          onClick={() => setStrokeOrderOpen(true)}
          title="Xem thứ tự nét bút"
          style={{
            position: "absolute",
            top: 6,
            left: 6,
            width: 20,
            height: 20,
            lineHeight: "18px",
            padding: 0,
            fontSize: 11,
            border: `1px solid ${COLORS.grid}`,
            borderRadius: "50%",
            background: COLORS.chipBg,
            color: COLORS.sealDark,
            cursor: "pointer",
          }}
        >
          ✍️
        </button>
      )}

      {mode === "view" && (
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
            border: `1px solid ${COLORS.grid}`,
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
          <div style={{ fontFamily: "KaiTi, 'STKaiti', 'Kaiti SC', 'Noto Serif SC', serif", fontSize: 26, color: COLORS.ink, textAlign: "center", marginBottom: 8 }}>
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
          <div style={{ fontFamily: "KaiTi, 'STKaiti', 'Kaiti SC', 'Noto Serif SC', serif", fontSize: 30, color: COLORS.ink }}>{b.char}</div>
          <div style={{ fontSize: 12.5, color: COLORS.sealDark, marginTop: 4 }}>{b.pinyin}</div>
          <div style={{ fontSize: 11.5, color: COLORS.inkSoft, marginTop: 2 }}>{b.meaning}</div>
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
                  ? "Đặt/cập nhật làm dữ liệu mặc định cho mọi người dùng mới"
                  : "Bấm để gỡ khỏi dữ liệu mặc định"
              }
              style={{
                marginTop: 6,
                fontSize: 10,
                padding: "3px 8px",
                borderRadius: 999,
                border: `1px solid ${!hasOverride && isOfficial ? COLORS.bamboo : COLORS.gold}`,
                background: !hasOverride && isOfficial ? "rgba(89,89,0,0.12)" : "rgba(89,89,0,0.06)",
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
        </>
      )}

      {strokeOrderOpen && <StrokeOrderModal char={b.char} onClose={() => setStrokeOrderOpen(false)} />}
    </div>
  );
}
