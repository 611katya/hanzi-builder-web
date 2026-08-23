import React, { useState, useEffect, useMemo, useCallback, useRef } from "react";
import * as XLSX from "xlsx";
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

/* ---------- Error boundary: surface crashes instead of silently blanking the UI ---------- */
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

export default function HanziBuilder({ userId }) {
  return (
    <ErrorBoundary>
      <HanziBuilderApp userId={userId} />
    </ErrorBoundary>
  );
}

function HanziBuilderApp({ userId }) {
  const [customBushou, setCustomBushou] = useState([]);
  const [customChars, setCustomChars] = useState([]);
  const [customWords, setCustomWords] = useState([]);
  const [deletedChars, setDeletedChars] = useState([]);
  const [needsReview, setNeedsReview] = useState([]);
  const [loaded, setLoaded] = useState(false);
  const [tab, setTab] = useState("play");

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
      if (!userId) return; // guest mode: keep in memory only, nothing to save
      const { error } = await supabase
        .from("custom_characters")
        .upsert(charToRow(entry, userId), { onConflict: "user_id,char" });
      if (error) console.error("Could not save character:", error);
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
    [...SEED_BUSHOU, ...customBushou].forEach((b) => map.set(b.char, b));
    return Array.from(map.values());
  }, [customBushou]);

  const characterList = useMemo(() => {
    const map = new Map();
    [...SEED_CHARACTERS, ...customChars].forEach((c) => map.set(c.char, c));
    deletedChars.forEach((ch) => map.delete(ch));
    return Array.from(map.values());
  }, [customChars, deletedChars]);

  const wordList = useMemo(() => {
    const map = new Map();
    [...SEED_WORDS, ...customWords].forEach((w) => map.set(w.word, w));
    return Array.from(map.values());
  }, [customWords]);

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
        <Header />
        <Tabs tab={tab} setTab={setTab} />

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
          />
        ) : (
          <RadicalsTab bushouList={bushouList} />
        )}
      </div>
    </div>
  );
}

/* ---------- Header ---------- */
function Header() {
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
      <div
        style={{
          fontFamily: "Calibri, 'Segoe UI', sans-serif",
          fontSize: 20,
          color: COLORS.sealDark,
          marginTop: 2,
        }}
      >
        Ghép Bộ Thủ Thành Chữ Hán
      </div>
      <div style={{ fontSize: 12.5, color: COLORS.inkSoft, marginTop: 6, letterSpacing: 0.3 }}>
        Build complete characters from their 部首 (bushou) components
      </div>
    </div>
  );
}

/* ---------- Tabs ---------- */
function Tabs({ tab, setTab }) {
  const items = [
    { id: "play", label: "Học · 学习" },
    { id: "add", label: "Thêm chữ · 添加" },
    { id: "radicals", label: "Bộ thủ · 部首" },
  ];
  return (
    <div
      style={{
        display: "flex",
        justifyContent: "center",
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

function PlayTab({ characterList, wordList, bushouList, findBushou, needsReview, onMarkNeedsReview, onClearNeedsReview }) {
  const [round, setRound] = useState(null); // { target, palette: [{id,char}] }
  const [selected, setSelected] = useState([]); // array of palette ids, in click order
  const [status, setStatus] = useState("playing"); // playing | correct | wrong | revealed
  const [score, setScore] = useState(0);
  const [streak, setStreak] = useState(0);
  const [selectedList, setSelectedList] = useState("Tất cả");
  const usedRef = useRef(new Set());

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
      sv: c.sv,
      lists: getLists(c),
      charGroups: buildCharGroups([c.char], characterList),
    }));
    const words = wordList.map((w) => ({
      key: w.word,
      display: w.word,
      pinyin: w.pinyin,
      meaning: w.meaning,
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

    const PALETTE_SIZE = 36;
    const neededSet = new Set(neededComponents);
    const distractPool = shuffle(bushouList.filter((b) => !neededSet.has(b.char)));
    const distractCount = Math.min(distractPool.length, Math.max(0, PALETTE_SIZE - neededComponents.length));
    const distractChips = distractPool.slice(0, distractCount).map((b) => ({ id: uid(), char: b.char, correct: false }));

    setRound({ target, palette: shuffle([...correctChips, ...distractChips]) });
    setSelected([]);
    setStatus("playing");
  }, [playable, bushouList]);

  useEffect(() => {
    usedRef.current = new Set();
    buildRound();
    // eslint-disable-next-line react-hooks/exhaustive-deps
  }, [selectedList, characterList.length]);

  const listPicker = (
    <div style={{ textAlign: "center", marginBottom: 16 }}>
      <select
        value={selectedList}
        onChange={(e) => setSelectedList(e.target.value)}
        style={{ ...selectStyle, width: 260, textAlign: "center", display: "inline-block" }}
      >
        <option value="Tất cả" style={{ background: COLORS.chipBg, color: COLORS.ink, fontWeight: 700 }}>Tất cả danh sách</option>
        <option value={REVIEW_LIST_VALUE} style={{ background: COLORS.chipBg, color: COLORS.ink, fontWeight: 700 }}>🔁 Cần ôn lại ({needsReview.length})</option>
        {allLists.map((l) => (
          <option key={l} value={l} style={{ background: COLORS.chipBg, color: COLORS.ink, fontWeight: 700 }}>
            {l}
          </option>
        ))}
      </select>
    </div>
  );

  if (!round) {
    return (
      <div>
        {listPicker}
        <div style={{ textAlign: "center", padding: 50, color: COLORS.inkSoft }}>
          {selectedList === REVIEW_LIST_VALUE
            ? "Danh sách ôn lại đang trống — nó chỉ chứa những chữ bạn đã dùng nút \"Xem đáp án\". Trả lời đúng một chữ sẽ tự động xóa nó khỏi danh sách này."
            : playable.length === 0
            ? 'Danh sách này chưa có chữ nào chơi được — có thể vì các chữ trong đó chưa có bộ thủ cấu thành. Ở tab "Thêm chữ", hãy dùng nút "🔍 Tự động điền" trước khi lưu để hệ thống tự nhận diện bộ thủ.'
            : 'Chưa có chữ nào trong kho dữ liệu. Hãy thêm chữ ở tab "Thêm chữ".'}
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
        <div style={{ fontFamily: "'Cormorant Garamond', serif", fontSize: 24, fontWeight: 600, color: COLORS.ink }}>
          {target.meaning}
        </div>
        <div style={{ display: "flex", justifyContent: "center", gap: 22, marginTop: 8, fontSize: 14.5, flexWrap: "wrap" }}>
          <span style={{ color: COLORS.sealDark }}>Pinyin: <strong>{target.pinyin}</strong></span>
          {target.sv && <span style={{ color: COLORS.bamboo }}>Hán Việt: <strong>{target.sv}</strong></span>}
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
                  <div style={{ fontSize: 12, color: COLORS.inkSoft, textAlign: "center" }}>chọn<br />bộ thủ bên dưới</div>
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
        {status === "correct" && <span style={{ color: COLORS.gold }}>✓ Chính xác! {target.display} ({target.pinyin}) — {target.meaning}</span>}
        {status === "wrong" && <span style={{ color: COLORS.error }}>✗ Chưa đúng. Đáp án đúng: {answerBreakdown}</span>}
        {status === "revealed" && (
          <span style={{ color: COLORS.gold }}>
            💡 Đáp án: {answerBreakdown} ({target.pinyin}) — {target.meaning}{target.sv ? `, Hán Việt: ${target.sv}` : ""}
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
            Undo - Chọn lại
          </button>
        )}
        {status === "playing" && (
          <button
            onClick={handleShowAnswer}
            className="ghost-btn"
            style={{ ...ghostBtnStyle, borderColor: COLORS.gold, color: COLORS.gold }}
          >
            💡 Xem đáp án
          </button>
        )}
        {status !== "playing" && (
          <button onClick={buildRound} className="seal-btn" style={sealBtnStyle}>
            Chữ tiếp theo →
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
}) {
  const [charInput, setCharInput] = useState("");
  const [meaning, setMeaning] = useState("");
  const [pinyin, setPinyin] = useState("");
  const [sv, setSv] = useState("");
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
    setLookupStatus("loading");
    try {
      const response = await fetch("/.netlify/functions/lookup-character", {
        method: "POST",
        headers: { "Content-Type": "application/json" },
        body: JSON.stringify({ char: target }),
      });
      if (!response.ok) {
        const errBody = await response.json().catch(() => ({}));
        throw new Error(errBody.error || `Lookup failed (${response.status})`);
      }
      const data = await response.json();
      const text = (data.content || []).map((b) => b.text || "").join("");
      const clean = text.replace(/```json|```/g, "").trim();
      const parsed = JSON.parse(clean);

      if (parsed.pinyin && (overwrite || !pinyin.trim())) setPinyin(parsed.pinyin);
      if (parsed.meaning && (overwrite || !meaning.trim())) setMeaning(parsed.meaning);
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
      setMessage({ type: "error", text: "Tra cứu tự động thất bại. Vui lòng nhập pinyin / nghĩa / Hán Việt thủ công." });
    }
  }

  function handleSubmit(e) {
    e.preventDefault();
    setMessage(null);
    try {
      if (!charInput.trim() || !meaning.trim() || !pinyin.trim() || !sv.trim()) {
        setMessage({ type: "error", text: "Vui lòng điền đầy đủ: chữ Hán, nghĩa, pinyin, âm Hán Việt." });
        return;
      }
      if (characterList.some((c) => c.char === charInput.trim())) {
        setMessage({ type: "error", text: `Chữ "${charInput.trim()}" đã có trong kho dữ liệu.` });
        return;
      }
      const listsToSave = selectedLists.length > 0 ? selectedLists : ["Chưa phân loại"];
      const trimmedChar = charInput.trim();
      onAddCharacter({
        char: trimmedChar,
        meaning: meaning.trim(),
        pinyin: pinyin.trim(),
        sv: sv.trim(),
        components: components,
        lists: listsToSave,
      });
      const playableNote =
        components.length > 0
          ? ""
          : ' (chưa có bộ thủ cấu thành nên sẽ không xuất hiện trong chế độ Chơi — dùng nút "🔍 Tự động điền" hoặc thêm bộ thủ thủ công trước khi lưu)';
      setMessage({ type: "success", text: `Đã thêm chữ "${trimmedChar}" vào danh sách "${listsToSave.join(", ")}"!${playableNote}` });
      resetForm();
    } catch (err) {
      console.error("Add character failed:", err);
      setMessage({ type: "error", text: `Có lỗi xảy ra: ${err && err.message ? err.message : "không rõ nguyên nhân"}. Vui lòng thử lại.` });
    }
  }

  return (
    <div>
      <div style={{ fontSize: 13, color: COLORS.inkSoft, marginBottom: 18, textAlign: "center" }}>
        Nhập một chữ Hán hoàn chỉnh cùng nghĩa, pinyin, âm Hán Việt, và xếp vào một danh sách (list) tuỳ chọn.
      </div>

      <div style={formCardStyle}>
        <FieldRow label="Chữ Hán hoàn chỉnh">
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
            {lookupStatus === "loading" ? "Đang tra…" : "🔍 Tự động điền"}
          </button>
        </FieldRow>
        <div style={{ fontSize: 11, color: COLORS.inkSoft, marginTop: -6, marginBottom: 14, paddingLeft: 204 }} className="autofill-hint">
          Gõ chữ Hán rồi rời khỏi ô để tự động điền pinyin, nghĩa, Hán Việt, và bộ thủ cấu thành — bạn vẫn có thể sửa lại thủ công.
        </div>

        <div style={{ borderTop: `1px dashed ${COLORS.grid}`, marginTop: 6, paddingTop: 14 }}>
          <div style={{ fontSize: 12.5, fontWeight: 600, color: COLORS.gold, marginBottom: 8, textTransform: "uppercase", letterSpacing: 0.8 }}>
            Bộ thủ cấu thành (theo thứ tự)
          </div>
          <div style={{ fontSize: 11, color: COLORS.inkSoft, marginBottom: 10 }}>
            Tự động điền ở trên sẽ gợi ý sẵn, hoặc bạn có thể thêm / xóa bộ thủ thủ công bên dưới. Cần bộ thủ để chữ này xuất hiện được ở chế độ Chơi.
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
              placeholder="gõ 1 bộ thủ, vd: 女"
              style={{ ...inputStyle, fontFamily: "KaiTi, 'STKaiti', 'Kaiti SC', 'Noto Serif SC', serif", width: 140 }}
              onKeyDown={(e) => {
                if (e.key === "Enter") {
                  e.preventDefault();
                  addComponentManually();
                }
              }}
            />
            <button type="button" onClick={addComponentManually} className="ghost-btn" style={ghostBtnStyle}>
              + Thêm bộ thủ
            </button>
          </div>

          {newCompDraft && (
            <div style={{ marginTop: 12, padding: 12, background: "rgba(169,130,47,0.08)", borderRadius: 8, border: `1px dashed ${COLORS.gold}` }}>
              <div style={{ fontSize: 12.5, marginBottom: 8 }}>
                Bộ thủ <strong style={{ fontFamily: "KaiTi, 'STKaiti', 'Kaiti SC', 'Noto Serif SC', serif", fontSize: 18 }}>{newCompDraft.char}</strong> chưa có trong kho — hãy điền thông tin:
              </div>
              <div style={{ display: "flex", gap: 8, flexWrap: "wrap" }}>
                <input value={ncPinyin} onChange={(e) => setNcPinyin(e.target.value)} placeholder="pinyin" style={{ ...inputStyle, width: 100 }} />
                <input value={ncMeaning} onChange={(e) => setNcMeaning(e.target.value)} placeholder="meaning" style={{ ...inputStyle, width: 140 }} />
                <input value={ncSv} onChange={(e) => setNcSv(e.target.value)} placeholder="âm Hán Việt" style={{ ...inputStyle, width: 120 }} />
                <input
                  value={ncStrokes}
                  onChange={(e) => setNcStrokes(e.target.value.replace(/[^0-9]/g, ""))}
                  placeholder="số nét (tùy chọn)"
                  style={{ ...inputStyle, width: 130 }}
                />
                <button type="button" onClick={confirmNewComponent} className="seal-btn" style={{ ...sealBtnStyle, padding: "8px 14px", fontSize: 13 }}>
                  Xác nhận
                </button>
              </div>
            </div>
          )}
        </div>

        <FieldRow label="Nghĩa (Meaning)">
          <input
            value={meaning}
            onChange={(e) => setMeaning(e.target.value)}
            onKeyDown={(e) => e.key === "Enter" && handleSubmit(e)}
            placeholder="good, well"
            style={inputStyle}
          />
        </FieldRow>

        <FieldRow label="Pinyin">
          <input
            value={pinyin}
            onChange={(e) => setPinyin(e.target.value)}
            onKeyDown={(e) => e.key === "Enter" && handleSubmit(e)}
            placeholder="hǎo"
            style={inputStyle}
          />
        </FieldRow>

        <FieldRow label="Âm Hán Việt (Sino-Vietnamese)">
          <input
            value={sv}
            onChange={(e) => setSv(e.target.value)}
            onKeyDown={(e) => e.key === "Enter" && handleSubmit(e)}
            placeholder="hảo"
            style={inputStyle}
          />
        </FieldRow>

        <FieldRow label="Danh sách (Lists)">
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
                placeholder="vd: HSK1, Gia đình, Bài 5… rồi Enter"
                list="existing-lists"
                style={inputStyle}
              />
              <button type="button" onClick={() => addList(listTypeahead)} className="ghost-btn" style={ghostBtnStyle}>
                + Thêm
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
          Một chữ có thể thuộc nhiều danh sách cùng lúc — gõ tên rồi nhấn Enter hoặc "+ Thêm" để thêm từng danh sách.
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
            Save
          </button>
        </div>
      </div>

      <AddWordPanel
        characterList={characterList}
        wordList={wordList}
        customWords={customWords}
        bushouList={bushouList}
        onAddCharacter={onAddCharacter}
        onAddBushou={onAddBushou}
        onAddWord={onAddWord}
        onDeleteWord={onDeleteWord}
      />

      <CharacterListPanel
        characterList={characterList}
        bushouList={bushouList}
        onDeleteCharacter={onDeleteCharacter}
        onUpdateCharacter={onUpdateCharacter}
        onAddBushou={onAddBushou}
      />
    </div>
  );
}

/* ---------- Add word: build a multi-character word from characters that
   already exist (or can be auto-filled on the spot), tag it with lists,
   and save it. Only shows/manages the current user's own custom words —
   the built-in seed words aren't editable here. ---------- */
function AddWordPanel({ characterList, wordList, customWords, bushouList, onAddCharacter, onAddBushou, onAddWord, onDeleteWord }) {
  const [expanded, setExpanded] = useState(false);
  const [wordInput, setWordInput] = useState("");
  const [pinyin, setPinyin] = useState("");
  const [meaning, setMeaning] = useState("");
  const [sv, setSv] = useState("");
  const [selectedLists, setSelectedLists] = useState([]);
  const [listTypeahead, setListTypeahead] = useState("");
  const [message, setMessage] = useState(null);
  const [charStatus, setCharStatus] = useState({}); // char -> "loading" | "error"
  const [wordLookupStatus, setWordLookupStatus] = useState("idle"); // idle | loading | error
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
    setCharStatus((prev) => ({ ...prev, [ch]: "loading" }));
    try {
      const response = await fetch("/.netlify/functions/lookup-character", {
        method: "POST",
        headers: { "Content-Type": "application/json" },
        body: JSON.stringify({ char: ch }),
      });
      if (!response.ok) throw new Error(`lookup failed (${response.status})`);
      const data = await response.json();
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
    setWordLookupStatus("loading");
    try {
      const response = await fetch("/.netlify/functions/lookup-word", {
        method: "POST",
        headers: { "Content-Type": "application/json" },
        body: JSON.stringify({ word }),
      });
      if (!response.ok) throw new Error(`lookup failed (${response.status})`);
      const data = await response.json();
      const text = (data.content || []).map((b) => b.text || "").join("");
      const clean = text.replace(/```json|```/g, "").trim();
      const parsed = JSON.parse(clean);
      if (parsed.pinyin) setPinyin(parsed.pinyin);
      if (parsed.meaning) setMeaning(parsed.meaning);
      if (parsed.sino_vietnamese) setSv(parsed.sino_vietnamese);
      if (!parsed.pinyin && !parsed.meaning) {
        setMessage({ type: "error", text: `Không tra được thông tin cho từ "${word}". Vui lòng nhập tay.` });
      }
      setWordLookupStatus("idle");
    } catch (err) {
      console.error("Word lookup failed:", err);
      setWordLookupStatus("error");
      setMessage({ type: "error", text: "Tra cứu từ thất bại. Vui lòng nhập tay pinyin / nghĩa / Hán Việt." });
    }
  }

  // One button (or one blur of the word field) does everything: looks up
  // pinyin/meaning/Hán Việt for the whole word, AND auto-fills any
  // character in it that doesn't have components yet.
  async function autoFillEverything() {
    if (chars.length < 2) return;
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

  function removeList(name) {
    setSelectedLists((prev) => prev.filter((l) => l !== name));
  }

  function handleSaveWord() {
    setMessage(null);
    const word = chars.join("");
    if (chars.length < 2) {
      setMessage({ type: "error", text: "Từ cần có ít nhất 2 chữ Hán." });
      return;
    }
    if (!allReady) {
      setMessage({ type: "error", text: "Mỗi chữ trong từ cần có bộ thủ trước — dùng nút \"Tự động điền\" bên dưới cho những chữ còn thiếu." });
      return;
    }
    if (!pinyin.trim() || !meaning.trim()) {
      setMessage({ type: "error", text: "Vui lòng điền pinyin và nghĩa của từ." });
      return;
    }
    if (wordList.some((w) => w.word === word)) {
      setMessage({ type: "error", text: `Từ "${word}" đã có trong kho dữ liệu.` });
      return;
    }
    const listsToSave = selectedLists.length > 0 ? selectedLists : ["Chưa phân loại"];
    onAddWord({ word, chars, pinyin: pinyin.trim(), meaning: meaning.trim(), sv: sv.trim(), lists: listsToSave });
    setMessage({ type: "success", text: `Đã thêm từ "${word}"!` });
    setWordInput("");
    setPinyin("");
    setMeaning("");
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
        {expanded ? "▲" : "▼"} Thêm từ nhiều chữ (vd: 你好)
      </button>

      {expanded && (
        <div style={{ marginTop: 12 }}>
          <div style={{ fontSize: 11.5, color: COLORS.inkSoft, marginBottom: 10 }}>
            Gõ một từ có từ 2 chữ Hán trở lên, rồi rời khỏi ô (hoặc bấm nút) để tự động điền pinyin, nghĩa, Hán
            Việt, và bộ thủ cho từng chữ còn thiếu — tất cả trong một bước.
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
              {wordLookupStatus === "loading" ? "Đang tra…" : "🔍 Tự động điền"}
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
                      <span style={{ fontSize: 11, color: COLORS.bamboo, fontWeight: 600 }}>✓ sẵn sàng</span>
                    ) : (
                      <>
                        <span style={{ fontSize: 11, color: COLORS.error }}>chưa có bộ thủ</span>
                        <button
                          type="button"
                          onClick={() => handleAutoFillChar(ch)}
                          disabled={status === "loading"}
                          className="ghost-btn"
                          style={{ ...ghostBtnStyle, padding: "3px 8px", fontSize: 10.5, borderColor: COLORS.gold, color: COLORS.gold }}
                        >
                          {status === "loading" ? "Đang tra…" : "🔍 Tự động điền"}
                        </button>
                      </>
                    )}
                    {status === "error" && <span style={{ fontSize: 10.5, color: COLORS.error }}>thất bại</span>}
                  </div>
                );
              })}
            </div>
          )}

          <FieldRow label="Pinyin cả từ">
            <input value={pinyin} onChange={(e) => setPinyin(e.target.value)} placeholder="tự động điền, hoặc nhập tay" style={inputStyle} />
          </FieldRow>
          <FieldRow label="Nghĩa cả từ">
            <input value={meaning} onChange={(e) => setMeaning(e.target.value)} placeholder="tự động điền, hoặc nhập tay" style={inputStyle} />
          </FieldRow>
          <FieldRow label="Hán Việt (tùy chọn)">
            <input value={sv} onChange={(e) => setSv(e.target.value)} placeholder="tự động điền, hoặc nhập tay" style={inputStyle} />
          </FieldRow>

          <FieldRow label="Danh sách (Lists)">
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
                  placeholder="vd: Thành ngữ… rồi Enter"
                  list="existing-word-lists"
                  style={inputStyle}
                />
                <button type="button" onClick={() => addList(listTypeahead)} className="ghost-btn" style={ghostBtnStyle}>
                  + Thêm
                </button>
              </div>
              <datalist id="existing-word-lists">
                {existingLists.map((l) => (
                  <option key={l} value={l} />
                ))}
              </datalist>
            </div>
          </FieldRow>

          {message && (
            <div style={{ marginTop: 8, fontSize: 12.5, fontWeight: 600, color: message.type === "error" ? COLORS.error : COLORS.bamboo }}>
              {message.text}
            </div>
          )}

          <div style={{ textAlign: "center", marginTop: 12 }}>
            <button type="button" onClick={handleSaveWord} className="seal-btn" style={{ ...sealBtnStyle, padding: "8px 20px", fontSize: 13 }}>
              Lưu từ
            </button>
          </div>

          {customWords && customWords.length > 0 && (
            <div style={{ marginTop: 16, paddingTop: 12, borderTop: `1px dashed ${COLORS.grid}` }}>
              <div style={{ fontSize: 11, color: COLORS.inkSoft, marginBottom: 8, textTransform: "uppercase", letterSpacing: 0.5, fontWeight: 700 }}>
                Từ của bạn ({customWords.length})
              </div>
              <div style={{ display: "flex", flexWrap: "wrap", gap: 8 }}>
                {customWords.map((w, idx) => (
                  <div
                    key={`${w.word}-${idx}`}
                    style={{
                      display: "flex",
                      alignItems: "center",
                      gap: 6,
                      padding: "5px 9px",
                      borderRadius: 6,
                      border: `1px solid ${COLORS.grid}`,
                      background: COLORS.card,
                      fontSize: 12,
                    }}
                  >
                    <span style={{ fontFamily: "KaiTi, 'STKaiti', 'Kaiti SC', 'Noto Serif SC', serif", fontSize: 16 }}>{w.word}</span>
                    <span style={{ color: COLORS.inkSoft }}>{w.meaning}</span>
                    <button
                      type="button"
                      onClick={() => onDeleteWord && onDeleteWord(w.word)}
                      style={{ background: "none", border: "none", color: COLORS.error, cursor: "pointer", fontSize: 12, padding: 0 }}
                      title="Xóa từ này"
                    >
                      ✕
                    </button>
                  </div>
                ))}
              </div>
            </div>
          )}
        </div>
      )}
    </div>
  );
}

/* ---------- List function: browse every character already in the database ---------- */
function CharacterListPanel({ characterList, bushouList, onDeleteCharacter, onUpdateCharacter, onAddBushou }) {
  const [query, setQuery] = useState("");
  const [listFilter, setListFilter] = useState("Tất cả");
  const [exportMessage, setExportMessage] = useState(null);

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
        setExportMessage({ type: "error", text: "Không có chữ nào để xuất." });
        return;
      }
      const rows = filtered.map((c) => ({
        "Chữ Hán": c.char,
        Pinyin: c.pinyin,
        "Nghĩa (Meaning)": c.meaning,
        "Âm Hán Việt": c.sv,
        "Danh sách": getLists(c).join(", "),
        "Bộ thủ cấu thành": (c.components || []).join(" + "),
      }));
      const worksheet = XLSX.utils.json_to_sheet(rows);
      worksheet["!cols"] = [
        { wch: 10 },
        { wch: 14 },
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
      setExportMessage({ type: "success", text: `Đã xuất ${filtered.length} chữ ra file Excel.` });
    } catch (err) {
      console.error("Export to Excel failed:", err);
      setExportMessage({ type: "error", text: "Xuất Excel thất bại. Vui lòng thử lại." });
    }
  }

  return (
    <div style={{ marginTop: 30 }}>
      <div style={{ borderTop: `1px dashed ${COLORS.grid}`, paddingTop: 22 }}>
        <div style={{ fontSize: 12.5, fontWeight: 600, color: COLORS.gold, marginBottom: 12, textTransform: "uppercase", letterSpacing: 0.8, textAlign: "center" }}>
          Danh sách chữ trong kho dữ liệu
        </div>

        <div style={{ display: "flex", justifyContent: "center", gap: 8, flexWrap: "wrap", marginBottom: 10 }}>
          <input
            value={query}
            onChange={(e) => setQuery(e.target.value)}
            placeholder="Tìm chữ theo Hán tự, pinyin, nghĩa, hoặc Hán Việt…"
            style={{ ...inputStyle, width: 300, textAlign: "center" }}
          />
          <select
            value={listFilter}
            onChange={(e) => setListFilter(e.target.value)}
            style={{ ...selectStyle, width: 170, flex: "none" }}
          >
            <option value="Tất cả" style={{ background: COLORS.chipBg, color: COLORS.ink, fontWeight: 700 }}>Tất cả danh sách</option>
            {allLists.map((l) => (
              <option key={l} value={l} style={{ background: COLORS.chipBg, color: COLORS.ink, fontWeight: 700 }}>
                {l}
              </option>
            ))}
          </select>
          <button
            type="button"
            onClick={handleExportExcel}
            className="seal-btn"
            style={{ ...sealBtnStyle, padding: "8px 16px", fontSize: 13, flex: "none" }}
          >
            ⬇ Xuất Excel
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
          {filtered.length} / {characterList.length} chữ
        </div>

        <div style={{ display: "grid", gridTemplateColumns: "repeat(auto-fill, minmax(170px, 1fr))", gap: 12 }}>
          {filtered.map((c) => (
            <CharacterCard
              key={c.char}
              c={c}
              bushouList={bushouList}
              findBushou={findBushou}
              onDeleteCharacter={onDeleteCharacter}
              onUpdateCharacter={onUpdateCharacter}
              onAddBushou={onAddBushou}
              allLists={allLists}
            />
          ))}
        </div>

        {filtered.length === 0 && (
          <div style={{ textAlign: "center", color: COLORS.inkSoft, fontSize: 13, padding: 30 }}>
            Không tìm thấy chữ nào phù hợp.
          </div>
        )}
      </div>
    </div>
  );
}

/* ---------- A single character card: view mode, edit mode, delete confirm ---------- */
function CharacterCard({ c, bushouList, findBushou, onDeleteCharacter, onUpdateCharacter, onAddBushou, allLists }) {
  const [mode, setMode] = useState("view"); // view | edit | confirmDelete
  const [meaning, setMeaning] = useState(c.meaning);
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

  function startEdit() {
    setMeaning(c.meaning);
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
          <label style={{ fontSize: 10, color: COLORS.inkSoft, display: "block", marginBottom: 2 }}>Nghĩa</label>
          <input value={meaning} onChange={(e) => setMeaning(e.target.value)} style={{ ...inputStyle, width: "100%", marginBottom: 6, fontSize: 12.5 }} />
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
            style={{
              fontSize: 10,
              color: COLORS.gold,
              fontWeight: 700,
              textTransform: "uppercase",
              letterSpacing: 0.4,
              marginBottom: 4,
              overflow: "hidden",
              textOverflow: "ellipsis",
              whiteSpace: "nowrap",
              paddingRight: 40,
            }}
            title={getLists(c).join(", ")}
          >
            {getLists(c).join(" · ")}
          </div>
          <div style={{ fontFamily: "KaiTi, 'STKaiti', 'Kaiti SC', 'Noto Serif SC', serif", fontSize: 32, color: COLORS.ink }}>{c.char}</div>
          <div style={{ fontSize: 12.5, color: COLORS.sealDark, marginTop: 4 }}>{c.pinyin}</div>
          <div style={{ fontSize: 11.5, color: COLORS.inkSoft, marginTop: 2 }}>{c.meaning}</div>
          <div style={{ fontSize: 11.5, color: COLORS.bamboo, marginTop: 2, fontWeight: 600 }}>HV: {c.sv}</div>
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
              <div style={{ fontSize: 11, color: COLORS.error, fontWeight: 600 }}>Xóa chữ "{c.char}"?</div>
              <div style={{ display: "flex", justifyContent: "center", gap: 6 }}>
                <button
                  type="button"
                  onClick={() => {
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
function RadicalsTab({ bushouList }) {
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
      <div style={{ textAlign: "center", marginBottom: 16 }}>
        <input
          value={query}
          onChange={(e) => setQuery(e.target.value)}
          placeholder="Tìm bộ thủ theo chữ, pinyin, nghĩa, hoặc Hán Việt…"
          style={{ ...inputStyle, width: 340, textAlign: "center" }}
        />
      </div>
      <div style={{ fontSize: 12.5, color: COLORS.inkSoft, textAlign: "center", marginBottom: 20 }}>
        {filtered.length} / {bushouList.length} bộ thủ · sắp xếp theo số nét
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
              <div
                key={b.char}
                style={{
                  background: COLORS.card,
                  border: `1px solid ${COLORS.grid}`,
                  borderRadius: 8,
                  padding: "12px 10px",
                  textAlign: "center",
                }}
              >
                <div style={{ fontFamily: "KaiTi, 'STKaiti', 'Kaiti SC', 'Noto Serif SC', serif", fontSize: 30, color: COLORS.ink }}>{b.char}</div>
                <div style={{ fontSize: 12.5, color: COLORS.sealDark, marginTop: 4 }}>{b.pinyin}</div>
                <div style={{ fontSize: 11.5, color: COLORS.inkSoft, marginTop: 2 }}>{b.meaning}</div>
                <div style={{ fontSize: 11.5, color: COLORS.bamboo, marginTop: 2, fontWeight: 600 }}>HV: {b.sv}</div>
              </div>
            ))}
          </div>
        </div>
      ))}
    </div>
  );
}
