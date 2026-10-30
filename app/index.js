const { app, ipcMain, dialog, globalShortcut } = require('electron');
const path = require('path');
const { LucidLog } = require('lucid-log');
const isDev = require('electron-is-dev');
const {
	isTrustedWebContents,
	isValidAccountId,
	isValidCount,
	isValidPartition,
	isValidText,
} = require('./security');

// Set app name for notifications BEFORE anything else
// Linux: use a stable unique identity for desktop integration (dock, Background Apps, tray)
if (process.platform === 'linux') {
	app.name = 'outlook-for-linux';
} else {
	app.name = 'Microsoft Outlook';
}
// Set desktop name to match the .desktop file for notification persistence
if (process.platform === 'linux') {
	app.setDesktopName('outlook-for-linux');
}

if (app.commandLine.hasSwitch('customUserDir')) {
	app.setPath('userData', app.commandLine.getSwitchValue('customUserDir'));
}

const { AppConfiguration } = require('./appConfiguration');
const appConfig = new AppConfiguration(app.getPath('userData'));

const config = appConfig.startupConfig;
config.appPath = path.join(__dirname, isDev ? '' : '../../');

const intuneEnabled = config.auth?.intune?.enabled;
if (config.tabbedMode && intuneEnabled) {
	config.tabbedMode = false;
	console.warn('Intune SSO is not supported with tabbed mode; disabling tabbed mode');
}

const logger = new LucidLog({
	levels: config.appLogLevels.split(','),
});

const notificationSounds = [
	{
		type: 'new-message',
		file: path.join(config.appPath, 'assets/sounds/new_message.wav'),
	},
];

// Notification sound player
/**
 * @type {NodeSoundPlayer}
 */
let player;
try {
	// eslint-disable-next-line no-unused-vars
	const { NodeSound } = require('node-sound');
	player = NodeSound.getDefaultPlayer();
} catch (e) {
	logger.info('No audio players found. Audio notifications might not work.');
}

const certificateModule = require('./certificate');
const notificationModule = require('./notification');
const gotTheLock = app.requestSingleInstanceLock();
const mainAppWindow = require('./mainAppWindow');
const QuickCompose = require('./quickCompose');

if (config.proxyServer)
	app.commandLine.appendSwitch('proxy-server', config.proxyServer);
app.commandLine.appendSwitch('disable-features', 'HardwareMediaKeyHandling');
app.commandLine.appendSwitch('enable-ntlm-v2', config.ntlmV2enabled);
app.commandLine.appendSwitch('try-supported-channel-layouts');

if (process.env.XDG_SESSION_TYPE === 'wayland') {
	logger.info('Running under Wayland, switching to PipeWire...');

	const features = app.commandLine.hasSwitch('enable-features')
		? app.commandLine.getSwitchValue('enable-features').split(',')
		: [];
	if (!features.includes('WebRTCPipeWireCapturer'))
		features.push('WebRTCPipeWireCapturer');

	app.commandLine.appendSwitch('enable-features', features.join(','));
	app.commandLine.appendSwitch('use-fake-ui-for-media-stream');
}

// Register protocol handlers
const protocols = ['msoutlook', 'mailto'];
protocols.forEach((protocol) => {
	if (!app.isDefaultProtocolClient(protocol, process.execPath)) {
		app.setAsDefaultProtocolClient(protocol, process.execPath);
		logger.info(`Registered as default protocol handler for: ${protocol}`);
	}
});

app.allowRendererProcessReuse = false;

if (!gotTheLock) {
	logger.info('App already running');
	app.quit();
} else {
	app.on('second-instance', mainAppWindow.onAppSecondInstance);
	app.on('open-url', (event, url) => {
		event.preventDefault();
		logger.info('open-url event received:', url);
		mainAppWindow.onAppSecondInstance(event, [url]);
	});
	app.on('ready', handleAppReady);
	app.on('quit', () => logger.debug('quit'));
	app.on('render-process-gone', onRenderProcessGone);
	app.on('will-quit', () => logger.debug('will-quit'));
	app.on('certificate-error', handleCertificateError);
	app.on('select-client-certificate', handleSelectClientCertificate);
	ipcMain.handle('getConfig', secureHandler(handleGetConfig));
	ipcMain.on('getWebauthnConfigSync', (event) => {
		event.returnValue = getWebauthnConfig();
	});
	ipcMain.handle('getZoomLevel', secureHandler(handleGetZoomLevel));
	ipcMain.handle('saveZoomLevel', secureHandler(handleSaveZoomLevel));
	ipcMain.handle('play-notification-sound', secureHandler(playNotificationSound));
	ipcMain.handle('set-badge-count', secureHandler(setBadgeCountHandler));
	ipcMain.handle('showEmailNotification', secureHandler(handleShowEmailNotification));
	ipcMain.handle('showReminderNotification', secureHandler(handleShowReminderNotification));
	ipcMain.handle('updateUnreadCount', secureHandler(handleUpdateUnreadCount));
	ipcMain.handle('updateReminderCount', secureHandler(handleUpdateReminderCount));
	ipcMain.handle('account-email-detected', secureHandler(handleAccountEmailDetected));
	ipcMain.handle('create-account', secureHandler(handleCreateAccount));
	ipcMain.handle('remove-account', secureHandler(handleRemoveAccount));
	ipcMain.handle('focus-account', secureHandler(handleFocusAccount));
	ipcMain.handle('toggle-auto-restore', secureHandler(handleToggleAutoRestore));
	ipcMain.handle('set-account-display-name', secureHandler(handleSetAccountDisplayName));
	ipcMain.handle('get-accounts', secureHandler(handleGetAccounts));

	// Tabbed mode IPC handlers
	ipcMain.handle('switch-tab', secureHandler(handleSwitchTab));
	ipcMain.handle('close-tab', secureHandler(handleCloseTab));

	// Quick Compose IPC handlers
	ipcMain.handle('open-quick-compose', secureHandler(handleOpenQuickCompose));
	ipcMain.handle('send-quick-compose', secureHandler(handleSendQuickCompose));
}

// Global reference to account manager (set by mainAppWindow)
let accountManager = null;

// Global reference to quick compose module
let quickCompose = null;

function setAccountManager(am) {
	accountManager = am;
	// Initialize quick compose after account manager is set
	quickCompose = new QuickCompose(accountManager, config);

	// Register global shortcut for quick compose (Ctrl+Shift+N)
	// Common email app shortcut
	globalShortcut.register('CommandOrControl+Shift+N', () => {
		if (quickCompose) {
			quickCompose.openDialog();
		}
	});
	logger.info('Registered global shortcut: Ctrl+Shift+N for Quick Compose');
}

// eslint-disable-next-line no-unused-vars
async function playNotificationSound(_event, options) {
	if (!options || typeof options !== 'object') return;
	const type = isValidText(options.type, 64) ? options.type : 'new-message';
	logger.debug(
		`Notification => Type: ${type}, Audio: ${options.audio}, Title: ${options.title}, Body: ${options.body}`,
	);
	if (!player || config.disableNotificationSound) return;

	const sound = notificationSounds.find((notificationSound) =>
		notificationSound.type === type,
	);
	if (sound) await player.play(sound.file);
}

function onRenderProcessGone() {
	logger.debug('render-process-gone');
	app.quit();
}

function onAppTerminated(signal) {
	if (signal === 'SIGTERM') {
		process.abort();
	} else {
		app.quit();
	}
}

function handleAppReady() {
	if (process.platform === 'linux' && config.webauthn) {
		require('./webauthn').initialize({ auth: { webauthn: { extraOrigins: parseOrigins(config.webauthnExtraOrigins) } } });
	}
	process.on('SIGTRAP', onAppTerminated);
	process.on('SIGINT', onAppTerminated);
	process.on('SIGTERM', onAppTerminated);
	//Just catch the error
	process.stdout.on('error', () => {});
	mainAppWindow.onAppReady(appConfig);
}

function parseOrigins(value) {
	return String(value || '').split(',').map((origin) => origin.trim()).filter(Boolean);
}

function getWebauthnConfig() {
	return {
		enabled: Boolean(config.webauthn),
		extraOrigins: parseOrigins(config.webauthnExtraOrigins),
	};
}

async function handleGetConfig() {
	return {
		appTitle: config.appTitle,
		disableNotifications: Boolean(config.disableNotifications),
		disableNotificationSound: Boolean(config.disableNotificationSound),
		partition: config.partition,
	};
}

async function handleGetZoomLevel(_event, name) {
	if (!isValidPartition(name)) return 0;
	const partition = getPartition(name) || {};
	return typeof partition.zoomLevel === 'number' ? partition.zoomLevel : 0;
}

async function handleSaveZoomLevel(_event, args) {
	if (
		!args ||
		typeof args !== 'object' ||
		!isValidPartition(args.partition) ||
		typeof args.zoomLevel !== 'number' ||
		!Number.isFinite(args.zoomLevel) ||
		args.zoomLevel < -9 ||
		args.zoomLevel > 9
	) {
		return;
	}
	let partition = getPartition(args.partition) || {};
	partition.name = args.partition;
	partition.zoomLevel = args.zoomLevel;
	savePartition(partition);
}

function getPartitions() {
	return appConfig.settingsStore.get('app.partitions') || [];
}

function getPartition(name) {
	const partitions = getPartitions();
	return partitions.filter((p) => {
		return p.name === name;
	})[0];
}

function savePartition(arg) {
	const partitions = getPartitions();
	const partitionIndex = partitions.findIndex((p) => {
		return p.name === arg.name;
	});

	if (partitionIndex >= 0) {
		partitions[partitionIndex] = arg;
	} else {
		partitions.push(arg);
	}
	appConfig.settingsStore.set('app.partitions', partitions);
}

function handleCertificateError() {
	const arg = {
		event: arguments[0],
		webContents: arguments[1],
		url: arguments[2],
		error: arguments[3],
		certificate: arguments[4],
		callback: arguments[5],
		config: config,
	};
	certificateModule.onAppCertificateError(arg, logger);
}

function handleSelectClientCertificate(
	event,
	webContents,
	url,
	list,
	callback,
) {
	event.preventDefault();

	logger.info(`Client certificate requested for URL: ${url}`);
	logger.info(`Available certificates: ${list.length}`);
	const certificate = certificateModule.selectClientCertificate({
		list,
		configuredSubject: config.clientCertSubject,
		dialog,
		logger,
	});
	if (certificate) logger.info(`Selected certificate: ${certificate.subjectName}`);
	callback(certificate || undefined);
}

async function setBadgeCountHandler(_event, count) {
	if (!isValidCount(count)) return;
	logger.debug(`Badge count set to '${count}'`);
	app.setBadgeCount(count);
}

async function handleShowEmailNotification(_event, notification) {
	if (
		!notification ||
		typeof notification !== 'object' ||
		!isValidText(notification.address, 320) ||
		!isValidText(notification.subject, 1000) ||
		(notification.body !== undefined && !isValidText(notification.body, 10000))
	) {
		return;
	}
	notificationModule.showEmailNotification(notification);
}

async function handleShowReminderNotification(_event, notification) {
	if (
		!notification ||
		typeof notification !== 'object' ||
		!isValidText(notification.subject || notification.text, 1000) ||
		(notification.timeUntil !== undefined && !isValidText(notification.timeUntil, 200))
	) {
		return;
	}
	notificationModule.showReminderNotification(notification);
}

async function handleUpdateUnreadCount(_event, data) {
	const count = typeof data === 'number' ? data : data?.count;
	const accountId = typeof data === 'object' && data ? data.accountId : null;
	if (!isValidCount(count)) return;

	if (accountId !== null) {
		if (!isValidAccountId(accountId) || !accountManager?.getAccount(accountId)) return;
		accountManager.updateUnreadCount(accountId, count);
		return;
	}
	notificationModule.updateBadgeFromUnreadCount(count);
}

async function handleUpdateReminderCount(_event, data) {
	const count = typeof data === 'number' ? data : data?.count;
	const accountId = typeof data === 'object' && data ? data.accountId : null;
	if (!isValidCount(count)) return;

	if (accountId !== null) {
		if (!isValidAccountId(accountId) || !accountManager?.getAccount(accountId)) return;
		accountManager.updateReminderCount(accountId, count);
		return;
	}
	notificationModule.updateBadgeFromReminderCount(count);
}

async function handleAccountEmailDetected(_event, data) {
	if (
		!data ||
		typeof data !== 'object' ||
		!isValidAccountId(data.accountId) ||
		!isValidText(data.email, 320) ||
		!accountManager?.getAccount(data.accountId)
	) {
		return;
	}
	accountManager.setAccountEmail(data.accountId, data.email);
}

async function handleCreateAccount(_event, options) {
	if (!accountManager || !options || typeof options !== 'object' || Array.isArray(options)) {
		return null;
	}
	const accountOptions = {};
	if (options.email !== undefined && isValidText(options.email, 320)) {
		accountOptions.email = options.email;
	}
	if (options.displayName !== undefined && isValidText(options.displayName, 200)) {
		accountOptions.displayName = options.displayName;
	}
	return accountManager.createAccount(accountOptions);
}

async function handleRemoveAccount(_event, accountId) {
	if (accountManager && isValidAccountId(accountId) && accountManager.getAccount(accountId)) {
		accountManager.removeAccount(accountId);
	}
}

async function handleFocusAccount(_event, accountId) {
	if (accountManager && isValidAccountId(accountId) && accountManager.getAccount(accountId)) {
		accountManager.focusAccount(accountId);
	}
}

async function handleToggleAutoRestore(_event, accountId) {
	if (accountManager && isValidAccountId(accountId) && accountManager.getAccount(accountId)) {
		accountManager.toggleAutoRestore(accountId);
	}
}

async function handleSetAccountDisplayName(_event, data) {
	if (
		!accountManager ||
		!data ||
		typeof data !== 'object' ||
		!isValidAccountId(data.accountId) ||
		!isValidText(data.displayName, 200) ||
		!accountManager.getAccount(data.accountId)
	) {
		return;
	}
	accountManager.setAccountDisplayName(data.accountId, data.displayName);
}

async function handleGetAccounts() {
	if (!accountManager) return [];
	return accountManager.getAllAccounts().map((account) => ({
		id: account.id,
		displayName: account.displayName,
		email: account.email,
		autoRestore: account.autoRestore !== false,
		unreadCount: account.unreadCount || 0,
		reminderCount: account.reminderCount || 0,
	}));
}

async function handleSwitchTab(_event, tabId) {
	if (
		accountManager?.tabManager &&
		isValidAccountId(tabId) &&
		accountManager.getAccount(tabId)
	) {
		accountManager.tabManager.switchTab(tabId);
	}
}

async function handleCloseTab(_event, tabId) {
	if (!accountManager || !isValidAccountId(tabId)) return;
	const account = accountManager.getAccount(tabId);
	if (!account) return;

	const label = account.email || account.displayName;
	const result = dialog.showMessageBoxSync({
		type: 'warning',
		buttons: ['Cancel', 'Close Tab'],
		defaultId: 0,
		cancelId: 0,
		title: 'Close Tab',
		message: `Close "${label}"?`,
		detail: 'This will remove the account from the tab bar. You can add it again later.',
	});

	if (result === 1) accountManager.removeAccount(tabId);
}

async function handleOpenQuickCompose(_event, accountId) {
	if (
		!quickCompose ||
		(accountId !== undefined &&
			accountId !== null &&
			(!isValidAccountId(accountId) || !accountManager?.getAccount(accountId)))
	) {
		return;
	}
	quickCompose.openDialog(accountId || null);
}

async function handleSendQuickCompose(_event, data) {
	if (
		!quickCompose ||
		!data ||
		typeof data !== 'object' ||
		!isValidText(data.to, 2000) ||
		!isValidText(data.subject || '', 1000) ||
		!isValidText(data.body || '', 100000) ||
		(data.accountId !== undefined &&
			(!isValidAccountId(data.accountId) || !accountManager?.getAccount(data.accountId)))
	) {
		return;
	}

	const mailtoLink =
		`mailto:${encodeURIComponent(data.to.trim())}` +
		`?subject=${encodeURIComponent(data.subject || '')}` +
		`&body=${encodeURIComponent(data.body || '')}`;
	const { shell } = require('electron');
	shell.openExternal(mailtoLink);
}

// Export setAccountManager for use by mainAppWindow
exports.setAccountManager = setAccountManager;
function secureHandler(handler) {
	return (event, ...args) => {
		if (!isTrustedWebContents(event.sender, config)) return;
		return handler(event, ...args);
	};
}
