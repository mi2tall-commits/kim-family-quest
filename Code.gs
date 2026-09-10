/**
 * 🏰 김가네 용돈 퀘스트 (FamilyQuest) - Google Apps Script 백엔드
 *
 * 📋 스프레드시트 시트 구조:
 *   - Users     : UserId | Name | Role | PIN | Points | TotalExp | Icon
 *   - Quests    : QuestId | Title | Category | Points | XP | Icon | Desc | Period
 *   - QuestLogs : LogId | QuestId | QuestTitle | UserId | SubmittedAt | Status | PhotoUrl | Memo | ParentComment | ApprovedBy
 *   - Rewards   : RewardId | Title | Cost | Category | Icon | Desc
 *   - Inventory : InvId | UserId | RewardId | RewardTitle | AcquiredAt | Status | Icon
 */

const SPREADSHEET_ID = "1GWeFz6kLzZSPz_XBhwO7cT69jSL2G5J4aWSC6EF1Hb4";
const FOLDER_ID      = "";
const FOLDER_NAME    = "FamilyQuest_Uploads";
const XP_PER_LEVEL   = 2500;

function doGet(e) {
  var action = (e && e.parameter && e.parameter.action) ? e.parameter.action : "";
  if (action === "initDatabase") {
    return ContentService.createTextOutput(JSON.stringify(initDatabase()))
      .setMimeType(ContentService.MimeType.JSON);
  }
  if (action === "getAppData" || (e && e.parameter && (e.parameter.api === "true" || e.parameter.format === "json"))) {
    return ContentService.createTextOutput(JSON.stringify(getAppData()))
      .setMimeType(ContentService.MimeType.JSON);
  }

  return HtmlService.createTemplateFromFile("index")
    .evaluate()
    .setTitle("김가네 용돈 퀘스트 | FamilyQuest")
    .addMetaTag("viewport", "width=device-width, initial-scale=1.0, maximum-scale=1.0, user-scalable=no, viewport-fit=cover")
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

function doPost(e) {
  try {
    var raw = e && e.postData ? e.postData.contents : "{}";
    var payload = {};
    try { payload = JSON.parse(raw); } catch(err) { payload = e.parameter || {}; }

    var action = payload.action || (e.parameter ? e.parameter.action : "");
    var data = payload.data || payload;
    var result = { success: false, error: "알 수 없는 요청입니다." };

    if (action === "initDatabase") {
      result = initDatabase();
    } else if (action === "getAppData") {
      result = getAppData();
    } else if (action === "completeQuest") {
      result = completeQuest(data);
    } else if (action === "approveQuest") {
      result = approveQuest(data);
    } else if (action === "rejectQuest") {
      result = rejectQuest(data);
    } else if (action === "redeemReward") {
      result = redeemReward(data);
    } else if (action === "useInventoryItem") {
      result = useInventoryItem(data);
    }

    return ContentService.createTextOutput(JSON.stringify(result))
      .setMimeType(ContentService.MimeType.JSON);
  } catch(err) {
    return ContentService.createTextOutput(JSON.stringify({ success: false, error: err.toString() }))
      .setMimeType(ContentService.MimeType.JSON);
  }
}

function getSS() {
  try { const a = SpreadsheetApp.getActiveSpreadsheet(); if (a) return a; } catch(e) {}
  return SpreadsheetApp.openById(SPREADSHEET_ID);
}

function getFolder() {
  if (FOLDER_ID) { try { return DriveApp.getFolderById(FOLDER_ID); } catch(e) {} }
  try {
    const it = DriveApp.getFoldersByName(FOLDER_NAME);
    if (it.hasNext()) return it.next();
    const f = DriveApp.createFolder(FOLDER_NAME);
    try { f.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW); } catch(e) {}
    return f;
  } catch(e) { return null; }
}

function sheetToObjects(sheet) {
  if (!sheet) return [];
  const lr = sheet.getLastRow();
  const lc = sheet.getLastColumn();
  if (lr < 2 || lc < 1) return [];
  const all = sheet.getRange(1, 1, lr, lc).getValues();
  const headers = all[0].map(function(h) { return String(h).trim(); });
  const result = [];
  for (var r = 1; r < all.length; r++) {
    var row = all[r];
    var hasData = row.some(function(c) { return String(c || "").trim().length > 0; });
    if (!hasData) continue;
    var obj = {};
    headers.forEach(function(h, i) {
      obj[h] = row[i];
      var clean = h.replace(/[^a-zA-Z0-9가-힣]/g, "").toLowerCase();
      if (clean) obj[clean] = row[i];
    });
    obj.__rowIndex = r + 1;
    result.push(obj);
  }
  return result;
}

function getField(obj) {
  if (!obj) return "";
  for (var i = 1; i < arguments.length; i++) {
    var k = arguments[i];
    if (obj[k] !== undefined && obj[k] !== null && String(obj[k]).trim() !== "") {
      return obj[k];
    }
    var clean = String(k).replace(/[^a-zA-Z0-9가-힣]/g, "").toLowerCase();
    if (obj[clean] !== undefined && obj[clean] !== null && String(obj[clean]).trim() !== "") {
      return obj[clean];
    }
  }
  return "";
}

function normCategory(raw) {
  var s = String(raw || "").trim().toLowerCase();
  if (s === "daily"   || s === "일일" || s === "??" || s === "?" || s.includes("?")) return "일일";
  if (s === "weekly"  || s === "주간") return "주간";
  if (s === "special" || s === "특별" || s === "bonus") return "특별";
  if (s.length > 0 && !s.includes("?")) return String(raw || "").trim();
  return "일일";
}

function getLevelTitle(lv) {
  if (lv >= 100) return "🎮 신화의 마스터 용사";
  if (lv >= 50)  return "🔮 대마법 아케이드 히어로";
  if (lv >= 30)  return "💎 플래티넘 챔피언";
  if (lv >= 20)  return "🛡️ 불꽃의 골드 가디언";
  if (lv >= 10)  return "⚔️ 용감한 실버 기사";
  if (lv >= 5)   return "🔥 열정의 프로 성실러";
  if (lv >= 3)   return "⭐ 든든한 퀘스트 챔피언";
  if (lv >= 2)   return "🌟 쑥쑥 자라는 모험가";
  return "🌱 새내기 견습 모험가";
}

function cleanIcon(icon, defaultIcon) {
  var s = String(icon || "").trim();
  if (!s || s === "?" || s === "??" || s.includes("?")) {
    return defaultIcon || "⭐";
  }
  return s;
}

function initDatabase() {
  try {
    var ss = getSS();

    // 1. Users 시트 초기화 (포인트 0, 다준용사)
    var usersSheet = ss.getSheetByName("Users");
    if (!usersSheet) usersSheet = ss.insertSheet("Users");
    usersSheet.clear();
    usersSheet.appendRow(["UserId", "Name", "Role", "PIN", "Points", "TotalExp", "Icon"]);
    usersSheet.appendRow(["child_1", "다준용사 (기사)", "CHILD", "", 0, 0, "🧒"]);
    usersSheet.appendRow(["mom", "엄마 (길드마스터)", "PARENT", "1234", 0, 0, "👩"]);
    usersSheet.appendRow(["dad", "아빠 (대마법사)", "PARENT", "1234", 0, 0, "👨"]);

    // 2. Quests 시트 초기화 (요청하신 샘플 2개)
    var questsSheet = ss.getSheetByName("Quests");
    if (!questsSheet) questsSheet = ss.insertSheet("Quests");
    questsSheet.clear();
    questsSheet.appendRow(["QuestId", "Title", "Category", "Points", "XP", "Icon", "Desc", "Period"]);
    questsSheet.appendRow(["Q_DISH_01", "설거지 & 싱크대 정리", "일일", 200, 200, "🍽️", "식사 후 그릇을 깨끗이 씻고 싱크대 물기 닦기", "일일"]);
    questsSheet.appendRow(["Q_DESK_01", "책상 정리", "일일", 300, 300, "🧹", "책상 위 먼지 털기 및 깔끔하게 정돈하기", "일일"]);

    // 3. Rewards 시트 초기화 (요청하신 샘플 2개)
    var rewardsSheet = ss.getSheetByName("Rewards");
    if (!rewardsSheet) rewardsSheet = ss.insertSheet("Rewards");
    rewardsSheet.clear();
    rewardsSheet.appendRow(["RewardId", "Title", "CostPoints", "Category", "Icon", "Description"]);
    rewardsSheet.appendRow(["R_GAME_30", "주말 게임 30분 쿠폰", 1500, "자유시간", "🎮", "주말에 게임할 수 있는 30분 자유 이용권"]);
    rewardsSheet.appendRow(["R_SNACK_01", "맛있는 아이스크림 1개", 2000, "간식", "🍦", "좋아하는 간식 또는 아이스크림 1개 교환권"]);

    // 4. QuestLogs 시트 초기화 (0점 유지를 위해 로그 삭제)
    var logsSheet = ss.getSheetByName("QuestLogs");
    if (!logsSheet) logsSheet = ss.insertSheet("QuestLogs");
    logsSheet.clear();
    logsSheet.appendRow(["LogId", "QuestId", "QuestTitle", "UserId", "SubmittedAt", "Status", "PhotoUrl", "Memo", "ParentComment", "ApprovedBy"]);

    // 5. Inventory 시트 초기화
    var invSheet = ss.getSheetByName("Inventory");
    if (!invSheet) invSheet = ss.insertSheet("Inventory");
    invSheet.clear();
    invSheet.appendRow(["InvId", "UserId", "RewardId", "RewardTitle", "AcquiredAt", "Status", "Icon"]);

    return { success: true, message: "✅ 김가네 스프레드시트 초기화 완료 (샘플 2개 & 0 포인트 세팅)", appData: getAppData() };
  } catch(e) {
    return { success: false, error: e.toString() };
  }
}

function getAppData() {
  try {
    var ss = getSS();

    var usersSheet   = ss.getSheetByName("Users");
    var questsSheet  = ss.getSheetByName("Quests");
    var logsSheet    = ss.getSheetByName("QuestLogs");
    var rewardsSheet = ss.getSheetByName("Rewards");
    var invSheet     = ss.getSheetByName("Inventory");

    var users   = sheetToObjects(usersSheet);
    var quests  = sheetToObjects(questsSheet);
    var logs    = sheetToObjects(logsSheet);
    var rewards = sheetToObjects(rewardsSheet);
    var invRows = sheetToObjects(invSheet);

    var profiles = users.map(function(u) {
      var rawName = String(getField(u, "name", "이름", "닉네임") || "").trim();
      var cleanName = rawName.replace(/\?/g, "").trim();
      if (!cleanName) cleanName = "다준용사 (기사)";
      var role = String(getField(u, "role", "역할", "구분") || "CHILD").trim().toUpperCase();
      var defIcon = role === "PARENT" ? (cleanName.includes("아빠") ? "👨" : "👩") : "🧒";
      return {
        id:       String(getField(u, "userId", "id", "아이디") || "user_" + u.__rowIndex).trim(),
        name:     cleanName,
        role:     role,
        pin:      String(getField(u, "pin", "비밀번호", "비번") || "").trim(),
        points:   Number(getField(u, "points", "포인트", "잔액") || 0) || 0,
        totalExp: Number(getField(u, "totalExp", "경험치", "총경험치", "xp", "points") || 0) || 0,
        icon:     cleanIcon(getField(u, "icon", "아이콘"), defIcon)
      };
    });

    var childProfile = profiles.find(function(p) { return p.role === "CHILD"; }) || {
      id: "child_1", name: "다준용사 (기사)", role: "CHILD", pin: "", points: 0, totalExp: 0, icon: "🧒"
    };

    var totalXp = childProfile.totalExp || 0;
    var currentBalance = childProfile.points || 0;

    var level = 1;
    if (totalXp > 0) level = Math.min(100, Math.max(1, Math.floor(totalXp / XP_PER_LEVEL) + 1));
    var levelTitle    = getLevelTitle(level);
    var lvProgress    = totalXp % XP_PER_LEVEL;
    var progressPct   = Math.min(100, Math.floor((lvProgress / XP_PER_LEVEL) * 100));

    var questList = quests.map(function(q) {
      var qId = String(getField(q, "questId", "id", "ID") || "Q-" + q.__rowIndex).trim();
      var qTitle = String(getField(q, "title", "퀘스트명", "제목", "퀘스트") || "").trim().replace(/\?/g, "").trim();
      if (!qTitle) qTitle = "퀘스트";
      var pts = Number(getField(q, "points", "포인트", "보상", "point", "Point") || 0) || 200;
      var icon = cleanIcon(getField(q, "icon", "아이콘"), "⭐");
      if (icon === "⭐") {
        if (qTitle.includes("설거지") || qTitle.includes("그릇")) icon = "🍽️";
        else if (qTitle.includes("책상") || qTitle.includes("정리")) icon = "🧹";
        else if (qTitle.includes("숙제") || qTitle.includes("공부")) icon = "✏️";
        else if (qTitle.includes("분리수거")) icon = "♻️";
        else if (qTitle.includes("신발")) icon = "👟";
        else if (qTitle.includes("양치")) icon = "🪥";
      }
      return {
        id:       qId,
        title:    qTitle,
        category: normCategory(getField(q, "category", "카테고리", "구분", "분류") || "일일"),
        period:   normCategory(getField(q, "period", "주기", "기간", "category") || "일일"),
        points:   pts,
        xp:       Number(getField(q, "xp", "경험치") || pts) || pts,
        icon:     icon,
        desc:     String(getField(q, "desc", "description", "설명", "내용") || "").trim().replace(/\?/g, "").trim()
      };
    }).filter(function(q) { return q.title.length > 0; });

    if (questList.length === 0) questList = getDefaultQuests();

    var questMap = {};
    questList.forEach(function(q) {
      if (q.id) questMap[q.id] = q;
      if (q.title) questMap[q.title] = q;
    });

    var profileMap = {
      "parent_1": "엄마 (길드마스터)",
      "parent_2": "아빠 (대마법사)",
      "mom": "엄마 (길드마스터)",
      "dad": "아빠 (대마법사)"
    };
    profiles.forEach(function(p) {
      if (p.id) profileMap[p.id] = p.name;
    });

    var tz = Session.getScriptTimeZone();
    var logList = logs.map(function(log) {
      var rawStatus = String(getField(log, "status", "상태") || "APPROVED").trim();
      var upperSt   = rawStatus.toUpperCase();
      var status    = upperSt === "PENDING" || rawStatus === "대기" ? "PENDING"
                    : upperSt === "REJECTED" || rawStatus === "반려" ? "REJECTED"
                    : "APPROVED";

      var dateRaw = getField(log, "submittedAt", "completedAt", "date", "일시", "날짜", "시간");
      var dateStr = "";
      if (dateRaw instanceof Date) {
        dateStr = Utilities.formatDate(dateRaw, tz, "yyyy-MM-dd HH:mm");
      } else if (dateRaw) {
        dateStr = String(dateRaw).trim();
      }

      var qId = String(getField(log, "questId", "id") || "").trim();
      var questObj = questMap[qId] || null;
      var qTitle = String(getField(log, "questTitle", "title", "퀘스트명", "제목") || (questObj ? questObj.title : "") || "퀘스트 완료").trim();
      var qPoints = Number(getField(log, "points", "포인트") || (questObj ? questObj.points : 300)) || 300;
      var qXp = Number(getField(log, "xp", "경험치") || (questObj ? questObj.xp : qPoints)) || qPoints;

      return {
        id:            String(getField(log, "logId", "id") || "LOG-" + log.__rowIndex).trim(),
        questId:       qId,
        title:         qTitle,
        userId:        String(getField(log, "userId") || "child_1").trim(),
        date:          dateStr,
        status:        status,
        photoUrl:      String(getField(log, "photoUrl", "photo_url", "사진") || "").trim(),
        memo:          String(getField(log, "memo", "소감", "메모") || "").trim(),
        parentComment: String(getField(log, "parentComment", "comment", "칭찬", "코멘트") || "").trim(),
        approvedBy:    String(getField(log, "approvedBy", "승인자") || "").trim(),
        points:        qPoints,
        xp:            qXp,
        rowIndex:      log.__rowIndex
      };
    }).reverse();

    var shopItems = rewards.map(function(r) {
      var title = String(getField(r, "title", "상품명", "보상", "name", "이름") || "").trim().replace(/\?/g, "").trim();
      var cost = Number(getField(r, "costPoints", "cost_points", "cost", "price", "points", "포인트") || 0) || 1000;
      var icon = cleanIcon(getField(r, "icon", "아이콘"), "🎁");
      if (icon === "🎁") {
        if (title.includes("게임") || title.includes("자유시간")) icon = "🎮";
        else if (title.includes("아이스크림") || title.includes("간식")) icon = "🍦";
        else if (title.includes("용돈") || title.includes("현금")) icon = "💰";
      }
      return {
        id:       String(getField(r, "rewardId", "id", "ID") || "R-" + r.__rowIndex).trim(),
        title:    title,
        cost:     cost,
        category: String(getField(r, "category", "카테고리", "분류") || "보상").trim().replace(/\?/g, "").trim() || "보상",
        icon:     icon,
        desc:     String(getField(r, "description", "desc", "설명", "내용") || "").trim().replace(/\?/g, "").trim()
      };
    }).filter(function(r) { return r.title.length > 0; });

    if (shopItems.length === 0) shopItems = getDefaultShopItems();

    var inventory = invRows.map(function(inv) {
      return {
        id:          String(getField(inv, "invId", "id") || "INV-" + inv.__rowIndex).trim(),
        userId:      String(getField(inv, "userId") || "child_1").trim(),
        rewardId:    String(getField(inv, "rewardId") || "").trim(),
        rewardTitle: String(getField(inv, "rewardTitle", "title", "보상명") || "").trim(),
        acquiredAt:  String(getField(inv, "acquiredAt", "date", "획득일") || "").trim(),
        status:      String(getField(inv, "status", "상태") || "ACTIVE").trim().toUpperCase(),
        icon:        cleanIcon(getField(inv, "icon", "아이콘"), "🎁"),
        rowIndex:    inv.__rowIndex
      };
    });

    var pendingReviews = logList.filter(function(l) { return l.status === "PENDING"; });

    var parentPins = { "mom": "1234", "dad": "1234" };
    profiles.filter(function(p) { return p.role === "PARENT"; }).forEach(function(p) {
      if (p.pin) parentPins[p.id] = p.pin;
      if (p.name.includes("엄마") || p.id === "mom") parentPins["mom"] = p.pin || "1234";
      if (p.name.includes("아빠") || p.id === "dad") parentPins["dad"] = p.pin || "1234";
    });

    return {
      profiles:       profiles,
      parentPins:     parentPins,
      child: {
        id:             childProfile.id,
        name:           childProfile.name,
        icon:           childProfile.icon,
        level:          level,
        title:          levelTitle,
        totalExp:       totalXp,
        currentBalance: currentBalance,
        targetItem:     "🎮 모험의 목표 보상",
        targetPoints:   100000,
        progressPct:    progressPct,
        lvProgress:     lvProgress,
        nextLvXp:       XP_PER_LEVEL
      },
      quests:         questList,
      shopItems:      shopItems,
      recentLogs:     logList,
      pendingReviews: pendingReviews,
      inventory:      inventory
    };
  } catch(e) {
    Logger.log("getAppData error: " + e);
    return getFallbackData();
  }
}

function completeQuest(data) {
  try {
    var ss = getSS();
    var tz = Session.getScriptTimeZone();
    var now = new Date();
    var nowStr = Utilities.formatDate(now, tz, "yyyy-MM-dd HH:mm:ss");
    var logId = "LOG_" + Utilities.formatDate(now, tz, "yyyyMMdd_HHmmss") + "_" + Math.floor(Math.random() * 9000 + 1000);

    var isParent = data.isParent || false;
    var status   = isParent ? "APPROVED" : "PENDING";

    var logsSheet = ss.getSheetByName("QuestLogs");
    if (!logsSheet) {
      logsSheet = ss.insertSheet("QuestLogs");
      logsSheet.appendRow(["LogId","QuestId","QuestTitle","UserId","SubmittedAt","Status","PhotoUrl","Memo","ParentComment","ApprovedBy"]);
    }

    logsSheet.appendRow([
      logId,
      data.questId || "CUSTOM",
      data.title || "퀘스트 완료",
      data.userId || "child_1",
      nowStr,
      status,
      data.photoUrl || "",
      data.memo || "",
      "",
      ""
    ]);

    if (isParent) {
      updateChildPoints(ss, data.userId || "child_1", Number(data.points) || 0);
    }

    return {
      success: true,
      message: isParent ? "🎉 퀘스트 승인 완료! +" + (data.points || 0) + " P" : "📤 퀘스트 제출 완료! 부모님 심의 대기 중입니다.",
      appData: getAppData()
    };
  } catch(err) {
    return { success: false, error: err.toString() };
  }
}

function approveQuest(data) {
  try {
    var ss = getSS();
    var sheet = ss.getSheetByName("QuestLogs");
    if (!sheet) return { success: false, error: "QuestLogs 시트 없음" };

    var rows = sheet.getRange(2, 1, Math.max(1, sheet.getLastRow() - 1), sheet.getLastColumn()).getValues();
    var found = -1;
    for (var i = 0; i < rows.length; i++) {
      if (String(rows[i][0]).trim() === String(data.logId).trim()) { found = i + 2; break; }
    }

    if (found !== -1) {
      sheet.getRange(found, 6).setValue("APPROVED");
      if (data.comment) sheet.getRange(found, 9).setValue(data.comment);
      sheet.getRange(found, 10).setValue(data.approvedBy || "부모님");

      var pts = Number(data.points) || 0;
      if (pts > 0) updateChildPoints(ss, data.userId || "child_1", pts);
    }

    return { success: true, message: "✅ 승인 완료! 포인트가 지급되었습니다.", appData: getAppData() };
  } catch(err) {
    return { success: false, error: err.toString() };
  }
}

function rejectQuest(data) {
  try {
    var ss = getSS();
    var sheet = ss.getSheetByName("QuestLogs");
    if (!sheet) return { success: false, error: "QuestLogs 시트 없음" };

    var rows = sheet.getRange(2, 1, Math.max(1, sheet.getLastRow() - 1), sheet.getLastColumn()).getValues();
    var found = -1;
    for (var i = 0; i < rows.length; i++) {
      if (String(rows[i][0]).trim() === String(data.logId).trim()) { found = i + 2; break; }
    }

    if (found !== -1) {
      sheet.getRange(found, 6).setValue("REJECTED");
      if (data.comment) sheet.getRange(found, 9).setValue(data.comment);
      sheet.getRange(found, 10).setValue(data.approvedBy || "부모님");
    }

    return { success: true, message: "❌ 반려 처리되었습니다.", appData: getAppData() };
  } catch(err) {
    return { success: false, error: err.toString() };
  }
}

function redeemReward(data) {
  try {
    var ss = getSS();
    var appData = getAppData();
    var cost = Number(data.cost) || 0;

    if (appData.child.currentBalance < cost) {
      return { success: false, error: "보유 포인트가 부족합니다!" };
    }

    var tz = Session.getScriptTimeZone();
    var now = Utilities.formatDate(new Date(), tz, "yyyy-MM-dd HH:mm:ss");
    var invId = "INV_" + Utilities.formatDate(new Date(), tz, "yyyyMMdd_HHmmss");

    var invSheet = ss.getSheetByName("Inventory");
    if (!invSheet) {
      invSheet = ss.insertSheet("Inventory");
      invSheet.appendRow(["InvId","UserId","RewardId","RewardTitle","AcquiredAt","Status","Icon"]);
    }
    invSheet.appendRow([invId, data.userId || "child_1", data.rewardId || "", data.title || "", now, "ACTIVE", data.icon || "🎁"]);

    updateChildPoints(ss, data.userId || "child_1", -cost);

    return { success: true, message: "🎁 [" + data.title + "] 획득! 인벤토리에 추가되었습니다.", appData: getAppData() };
  } catch(err) {
    return { success: false, error: err.toString() };
  }
}

function useInventoryItem(data) {
  try {
    var ss = getSS();
    var sheet = ss.getSheetByName("Inventory");
    if (!sheet) return { success: false, error: "Inventory 시트 없음" };

    var rows = sheet.getRange(2, 1, Math.max(1, sheet.getLastRow() - 1), sheet.getLastColumn()).getValues();
    var found = -1;
    for (var i = 0; i < rows.length; i++) {
      if (String(rows[i][0]).trim() === String(data.invId).trim()) { found = i + 2; break; }
    }

    if (found !== -1) {
      sheet.getRange(found, 6).setValue("USED");
    }

    return { success: true, message: "✅ 아이템을 사용했습니다!", appData: getAppData() };
  } catch(err) {
    return { success: false, error: err.toString() };
  }
}

function updateChildPoints(ss, userId, delta) {
  var usersSheet = ss.getSheetByName("Users");
  if (!usersSheet || usersSheet.getLastRow() < 2) return;

  var data = usersSheet.getRange(1, 1, usersSheet.getLastRow(), usersSheet.getLastColumn()).getValues();
  var headers = data[0].map(function(h) { return String(h).trim().toLowerCase(); });
  var idCol  = headers.indexOf("userid");
  var ptCol  = headers.indexOf("points");
  var xpCol  = headers.indexOf("totalexp");
  if (idCol === -1) idCol = 0;
  if (ptCol === -1) ptCol = 4;
  if (xpCol === -1) xpCol = 5;

  for (var r = 1; r < data.length; r++) {
    var rowId = String(data[r][idCol]).trim();
    var role  = String(data[r][headers.indexOf("role") !== -1 ? headers.indexOf("role") : 2] || "").toUpperCase();
    if (rowId === userId || role === "CHILD") {
      var curPts = Number(data[r][ptCol]) || 0;
      var curXp  = Number(data[r][xpCol]) || 0;
      usersSheet.getRange(r + 1, ptCol + 1).setValue(curPts + delta);
      if (delta > 0) usersSheet.getRange(r + 1, xpCol + 1).setValue(curXp + delta);
      break;
    }
  }
}

function getDefaultQuests() {
  return [
    { id: "Q_DISH_01", title: "설거지 & 싱크대 정리", category: "일일", period: "일일", points: 200, xp: 200, icon: "🍽️", desc: "식사 후 그릇을 깨끗이 씻고 싱크대 물기 닦기" },
    { id: "Q_DESK_01", title: "책상 정리", category: "일일", period: "일일", points: 300, xp: 300, icon: "🧹", desc: "책상 위 먼지 털기 및 깔끔하게 정돈하기" }
  ];
}

function getDefaultShopItems() {
  return [
    { id: "R_GAME_30", title: "주말 게임 30분 쿠폰", cost: 1500, category: "자유시간", icon: "🎮", desc: "주말에 게임할 수 있는 30분 자유 이용권" },
    { id: "R_SNACK_01", title: "맛있는 아이스크림 1개", cost: 2000, category: "간식", icon: "🍦", desc: "좋아하는 간식 또는 아이스크림 1개 교환권" }
  ];
}

function getFallbackData() {
  return {
    profiles: [
      { id: "child_1", name: "다준용사 (기사)", role: "CHILD", pin: "", points: 0, totalExp: 0, icon: "🧒" },
      { id: "mom", name: "엄마 (길드마스터)", role: "PARENT", pin: "1234", points: 0, totalExp: 0, icon: "👩" },
      { id: "dad", name: "아빠 (대마법사)", role: "PARENT", pin: "1234", points: 0, totalExp: 0, icon: "👨" }
    ],
    parentPins: { mom: "1234", dad: "1234" },
    child: {
      id: "child_1", name: "다준용사 (기사)", icon: "🧒",
      level: 1, title: "🌱 새내기 견습 모험가",
      totalExp: 0, currentBalance: 0,
      targetItem: "🎮 모험의 목표 보상", targetPoints: 100000,
      progressPct: 0, lvProgress: 0, nextLvXp: 2500
    },
    quests: getDefaultQuests(),
    shopItems: getDefaultShopItems(),
    recentLogs: [],
    pendingReviews: [],
    inventory: []
  };
}
