/*******************************************************
 * اتصال Google Drive — کتابخانه محتوا
 * فهرست‌برداری صفحه‌ای (بدون بارگذاری کل Drive) + کنترل دسترسی سطح پوشه
 *******************************************************/

function listChildren_(folderId, foldersOnly) {
  const q = ["'" + folderId + "' in parents", 'trashed = false'];
  if (foldersOnly) q.push("mimeType = 'application/vnd.google-apps.folder'");
  const files = [];
  let pageToken = null;
  do {
    const res = Drive.Files.list({
      q: q.join(' and '),
      pageSize: 200,
      pageToken: pageToken || undefined,
      fields: 'nextPageToken,files(id,name,mimeType)',
      supportsAllDrives: true,
      includeItemsFromAllDrives: true
    });
    (res.files || []).forEach(f => files.push(f));
    pageToken = res.nextPageToken;
  } while (pageToken && files.length < 1000);
  return files;
}

/** نگاشت کلید تیم → پوشه Drive (کَش ۱۰ دقیقه‌ای) */
function findTeamFolders_(user) {
  const cacheKey = 'team_folders_v1';
  let folders = cacheGet_(cacheKey);
  if (!folders) {
    try {
      folders = listChildren_(CONFIG.ROOT_FOLDER_ID, true);
    } catch (e) {
      throw new Error('دسترسی به پوشه اصلی Drive ممکن نشد. ROOT_FOLDER_ID و اشتراک‌گذاری پوشه را بررسی کنید. جزئیات: ' + e.message);
    }
    cachePut_(cacheKey, folders, 600);
  }
  return TEAM_CONFIG
    .filter(team => allowedTeams_(user).indexOf(team.key) !== -1)
    .map(team => {
      const folder = folders.find(f =>
        team.aliases.some(alias => String(f.name).toLowerCase() === String(alias).toLowerCase())
      );
      return folder ? {
        id: folder.id, key: team.key, name: team.name,
        color: team.color, icon: team.icon, mimeType: folder.mimeType
      } : null;
    })
    .filter(Boolean);
}

function getTeamFolderOrThrow_(user, teamKey) {
  assertTeamAccess_(user, teamKey);
  const team = findTeamFolders_(user).find(t => t.key === teamKey);
  if (!team) throw new Error('پوشه تیم «' + (teamConfigByKey_(teamKey) || {}).name + '» در Drive پیدا نشد.');
  return team;
}

function getDriveName_(id) {
  const cacheKey = 'fname_' + id;
  const cached = cacheGet_(cacheKey);
  if (cached) return cached;
  const meta = Drive.Files.get(id, { fields: 'name,mimeType', supportsAllDrives: true });
  cachePut_(cacheKey, meta.name, 600);
  return meta.name;
}

/**
 * صفحه‌ای از محتوای یک پوشه برمی‌گرداند.
 * دسترسی: targetId باید داخل پوشه تیم مجاز کاربر باشد.
 */
function getLibraryPage(folderId, pageToken, searchText, typeFilter, teamKey) {
  const user = getUserContext();
  const team = getTeamFolderOrThrow_(user, teamKey);
  const targetId = folderId || team.id;
  if (!isWithinTeam_(targetId, team.id)) {
    throw new Error('دسترسی به این پوشه مجاز نیست.');
  }

  const q = ["'" + targetId + "' in parents", 'trashed = false'];
  if (searchText) q.push("name contains '" + escapeQuery_(searchText) + "'");
  if (typeFilter === 'folder') q.push("mimeType = 'application/vnd.google-apps.folder'");
  if (typeFilter === 'file') q.push("mimeType != 'application/vnd.google-apps.folder'");

  // شمارش سریع نتایج صفحه نه کل نتایج (محدودیت Drive API برای count وجود ندارد)
  const response = Drive.Files.list({
    q: q.join(' and '),
    pageSize: CONFIG.PAGE_SIZE,
    pageToken: pageToken || undefined,
    orderBy: 'folder,name',
    fields: 'nextPageToken,files(id,name,mimeType,modifiedTime,size,webViewLink,thumbnailLink,iconLink,parents,description)',
    supportsAllDrives: true,
    includeItemsFromAllDrives: true
  });

  const myProgress = getMyProgressMap_(user.email);
  const quizMap = getQuizRefMap_();

  const items = (response.files || []).map(f => {
    const isFolder = f.mimeType === 'application/vnd.google-apps.folder';
    const quizId = quizMap[f.id] || '';
    return {
      id: f.id,
      name: f.name,
      mimeType: f.mimeType,
      kind: isFolder ? 'folder' : 'file',
      modified: f.modifiedTime || '',
      size: formatSize_(f.size),
      url: f.webViewLink || ('https://drive.google.com/open?id=' + f.id),
      previewUrl: isFolder ? '' : 'https://drive.google.com/file/d/' + f.id + '/preview',
      thumbnail: f.thumbnailLink || '',
      icon: fileIcon_(f.mimeType),
      quiz: quizId ? { available: true, quizId: quizId } : { available: false },
      status: myProgress['drive:' + f.id] ? 'completed' : 'not_started'
    };
  });

  return ok_({
    items: items,
    nextPageToken: response.nextPageToken || '',
    hasMore: !!response.nextPageToken,
    pageSize: CONFIG.PAGE_SIZE,
    folderId: targetId,
    folderName: getDriveName_(targetId),
    teamId: team.id
  });
}

/** جزئیات یک آیتم + آزمون متصل + وضعیت کاربر */
function getItemDetail(teamKey, itemId) {
  const user = getUserContext();
  const team = getTeamFolderOrThrow_(user, teamKey);
  if (!isWithinTeam_(itemId, team.id)) throw new Error('دسترسی به این آیتم مجاز نیست.');

  const meta = Drive.Files.get(itemId, {
    fields: 'id,name,mimeType,modifiedTime,size,webViewLink,thumbnailLink,description',
    supportsAllDrives: true
  });
  const quizMap = getQuizRefMap_();
  const quizId = quizMap[itemId] || '';
  const progress = getMyProgressMap_(user.email);
  return ok_({
    item: {
      id: meta.id, name: meta.name, mimeType: meta.mimeType,
      kind: meta.mimeType === 'application/vnd.google-apps.folder' ? 'folder' : 'file',
      modified: meta.modifiedTime || '', size: formatSize_(meta.size),
      url: meta.webViewLink || '', thumbnail: meta.thumbnailLink || '',
      description: meta.description || '', icon: fileIcon_(meta.mimeType),
      status: progress['drive:' + meta.id] ? 'completed' : 'not_started'
    },
    quiz: quizId ? getQuizMetaById_(quizId) : { available: false }
  });
}

/** ثبت تکمیل مشاهده یک فایل Drive (جدا از مسیر یادگیری) */
function markDriveItemComplete(teamKey, itemId) {
  const user = getUserContext();
  const team = getTeamFolderOrThrow_(user, teamKey);
  if (!isWithinTeam_(itemId, team.id)) throw new Error('دسترسی به این آیتم مجاز نیست.');
  markProgress_(user.email, 'drive:' + itemId, 100);
  logActivity_(user.email, 'item_complete', itemId);
  return ok_({});
}
