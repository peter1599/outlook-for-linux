
const { BrowserWindow, nativeTheme } = require("electron");
const isDarkMode = nativeTheme.shouldUseDarkColors;
const path = require("path");
const Menus = require("../menus");
const { LucidLog } = require("lucid-log");
const TrayIconChooser = require("../browser/tools/trayIconChooser");
// eslint-disable-next-line no-unused-vars
const { AppConfiguration } = require("../appConfiguration");
const notificationModule = require("../notification");
const QuickCompose = require("../quickCompose");
const { isSafeExternalUrl, isExternalNavigation } = require("../security");

/**
 * @type {TrayIconChooser}
 */
let iconChooser;

// NOTE: isControlPressed removed - no longer needed in multi-account mode

/**
 * @type {LucidLog}
 */
let logger;

// NOTE: aboutBlankRequestCount removed - no longer needed in multi-account mode
let config;

// NOTE: window variable removed - no longer a single main window in multi-account mode

/**
 * @type {AccountManager}
 */
let accountManager;
let intune;

/**
 * @param {AppConfiguration} mainConfig
 */
exports.onAppReady = async function onAppReady(mainConfig) {
	const AccountManager = require("../accountManager");
	const mainApp = require("../index");

	config = mainConfig.startupConfig;
	logger = new LucidLog({
		levels: config.appLogLevels.split(","),
	});
	if (config.auth?.intune?.enabled) {
		intune = require("../intune");
		await intune.initSso(config.auth.intune.user || "");
	}
	iconChooser = new TrayIconChooser(mainConfig.startupConfig);

	// Create menus instance (will be shared across all accounts)
	const menus = new Menus(null, config, iconChooser.getFile(), mainConfig);

	// Create account manager
	const accMgr = new AccountManager(
		config,
		iconChooser.getFile(),
		menus,
		mainConfig,
	);

	// Create quick compose instance
	const qCompose = new QuickCompose(accMgr, config);

	// Set as global reference for IPC handlers
	mainApp.setAccountManager(accMgr);

	// Initialize notification module (will be called by each account window)
	notificationModule.init(null, iconChooser.getFile(), menus);

	// Wire up quick compose to menus
	if (menus.tray) {
		menus.tray.setQuickCompose(qCompose);
	}
	if (menus.setQuickCompose) {
		menus.setQuickCompose(qCompose);
	}

	// Set as global reference for IPC handlers
	mainApp.setAccountManager(accMgr);

	// Initialize notification module (will be called by each account window)
	notificationModule.init(null, iconChooser.getFile(), menus);

	// Store references for other functions
	// Store in module-level variable for access from event handlers (which lose 'this' context)
	accountManager = accMgr;
	this.accountManager = accMgr;
	this.quickCompose = qCompose;
	this.menus = menus;

	// Restore saved accounts (or create default if none exist)
	// If multiple accounts exist, show chooser first
	const accounts = accMgr.getAllAccounts();
	if (accounts.length === 0) {
		// No accounts, create default one
		await accountManager.restoreAccounts();
	} else if (accounts.length === 1) {
		// Single account, just restore it
		await accountManager.restoreAccounts();
	} else {
		// Multiple accounts - show chooser
		await showStartupAccountChooser(accountManager);
	}

	// Note: addEventHandlers() removed - no longer a single main window

	// Handle command line args (e.g., mailto links)
	const result = processArgs(process.argv);
	if (result && result.isMailto) {
		// Open mailto in new compose window
		await openComposeWindow(result.url, accountManager);
	}
};

let allowFurtherRequests = true;

exports.onAppSecondInstance = async function onAppSecondInstance(event, args) {
	logger.debug("second-instance started");
	event.preventDefault();

	const result = processArgs(args);
	if (result && allowFurtherRequests) {
		allowFurtherRequests = false;
		setTimeout(() => {
			allowFurtherRequests = true;
		}, 5000);

		if (result.isMailto) {
			// Open mailto links in new compose window
			await openComposeWindow(result.url, accountManager);
		} else {
			// For other URLs, load in the first available account window
			const accounts = accountManager?.getAllAccounts() || [];
			const firstAccount = accounts.find(
				(a) => a.window && !a.window.isDestroyed(),
			);
			if (firstAccount) {
				firstAccount.window.loadURL(result.url, {
					userAgent: config.chromeUserAgent,
				});
				firstAccount.window.show();
				firstAccount.window.focus();
			}
		}
	}

	// Show all account windows (but not after mailto - user may have cancelled)
	if (!result?.isMailto) {
		accountManager?.showAllWindows();
	}
};

// Export functions for external access
exports.getAccountManager = function () {
	return this.accountManager;
};

exports.getMenus = function () {
	return this.menus;
};

// NOTE: Old single-window functions removed - now handled per-account in AccountManager
// - applyAppConfiguration (no longer needed)
// - restoreWindow (no longer needed)

function processArgs(args) {
	var regHttps =
		/^https:\/\/outlook.microsoft.com\/l\/(meetup-join|channel)\//g;
	var regMS = /^msoutlook:\/l\/(meetup-join|channel)\//g;
	var regMailto = /^mailto:/i;
	logger.debug("processArgs:", args);
	for (const arg of args) {
		if (regHttps.test(arg)) {
			logger.debug("A url argument received with https protocol");
			// Note: window.show() removed - window is null in multi-account mode
			return { url: arg, isMailto: false };
		}
		if (regMS.test(arg)) {
			logger.debug("A url argument received with msoutlook protocol");
			// Note: window.show() removed - window is null in multi-account mode
			return {
				url: config.url + arg.substring(8, arg.length),
				isMailto: false,
			};
		}
		if (regMailto.test(arg)) {
			logger.debug("A mailto argument received");
			// Note: window.show() removed - window is null in multi-account mode
			// Convert mailto: URL to Outlook compose URL
			return { url: convertMailtoToOutlookURL(arg), isMailto: true };
		}
	}
}

function showAccountSelectionDialog(accountManager) {
	return new Promise((resolve) => {
		const { BrowserWindow, ipcMain } = require('electron');
		const accounts = accountManager.getAllAccounts();
		if (accounts.length === 0) {
			resolve(null);
			return;
		}

		const isDarkMode = nativeTheme.shouldUseDarkColors;
		const height = Math.min(140 + accounts.length * 45, 500);
		const selectDialog = new BrowserWindow({
			width: 400,
			height,
			resizable: false,
			modal: false,
			show: false,
			autoHideMenuBar: true,
			title: 'Select Account',
			backgroundColor: isDarkMode ? '#1e1e1e' : '#ffffff',
			webPreferences: {
				preload: path.join(__dirname, '..', 'dialogs', 'preload.js'),
				nodeIntegration: false,
				contextIsolation: true,
				sandbox: true,
				additionalArguments: [
					'--dialog-type=mailto-account',
					`--dialog-data=${encodeURIComponent(JSON.stringify({
						isDarkMode,
						accounts: accounts.map(({ id, displayName }) => ({ id, displayName })),
					}))}`,
				],
			},
		});

		selectDialog.loadFile(path.join(__dirname, '..', 'dialogs', 'dialog.html'));

		const handler = (event, accountId) => {
			if (
				event.sender !== selectDialog.webContents ||
				!accounts.some((account) => account.id === accountId)
			) {
				return;
			}
			resolve(accountId);
			ipcMain.removeListener('mailto-account-selected', handler);
			selectDialog.close();
		};

		ipcMain.on('mailto-account-selected', handler);
		selectDialog.on('closed', () => {
			ipcMain.removeListener('mailto-account-selected', handler);
			resolve(null);
		});
		selectDialog.once('ready-to-show', () => {
			selectDialog.show();
			selectDialog.focus();
		});
	});
}

async function showStartupAccountChooser(accountManager) {
	const { BrowserWindow, ipcMain } = require('electron');
	const accounts = accountManager.getAllAccounts();
	const isDarkMode = nativeTheme.shouldUseDarkColors;
	const height = Math.min(180 + accounts.length * 50, 500);
	const selectDialog = new BrowserWindow({
		width: 420,
		height,
		resizable: false,
		modal: false,
		show: false,
		autoHideMenuBar: true,
		title: 'Select Accounts',
		backgroundColor: isDarkMode ? '#1e1e1e' : '#ffffff',
		webPreferences: {
			preload: path.join(__dirname, '..', 'dialogs', 'preload.js'),
			nodeIntegration: false,
			contextIsolation: true,
			sandbox: true,
			additionalArguments: [
				'--dialog-type=startup-account',
				`--dialog-data=${encodeURIComponent(JSON.stringify({
					isDarkMode,
					accounts: accounts.map(({ id, displayName, autoRestore }) => ({
						id,
						displayName,
						autoRestore,
					})),
				}))}`,
			],
		},
	});

	selectDialog.loadFile(path.join(__dirname, '..', 'dialogs', 'dialog.html'));

	const handler = (event, accountIds) => {
		if (event.sender !== selectDialog.webContents || !Array.isArray(accountIds)) {
			return;
		}
		const validIds = accountIds.filter((id) =>
			accounts.some((account) => account.id === id),
		);
		validIds.forEach((id) => {
			const account = accountManager.getAccount(id);
			if (account) accountManager.createAccountWindow(account);
		});
		ipcMain.removeListener('startup-account-selection', handler);
		selectDialog.close();
	};

	ipcMain.on('startup-account-selection', handler);
	selectDialog.on('closed', () => {
		ipcMain.removeListener('startup-account-selection', handler);
	});
	selectDialog.once('ready-to-show', () => {
		selectDialog.show();
		selectDialog.focus();
	});
}

/**
 * Open a compose window with the given URL in the selected account's session
 * @param {string} url - Outlook compose URL
 * @param {AccountManager} [accountManager] - Account manager for account selection
 */
async function openComposeWindow(url, accountManager) {
	let targetPartition = config.partition;
	let targetAccountId = null;

	// If account manager provided, determine which account to use
	if (accountManager) {
		const accounts = accountManager.getAllAccounts();
		if (accounts.length === 0) {
			logger.debug("No accounts available, using default partition");
		} else if (accounts.length === 1) {
			// Single account - use it directly
			targetPartition = accounts[0].partition;
			targetAccountId = accounts[0].id;
		} else {
			// Multiple accounts - show selection dialog
			targetAccountId = await showAccountSelectionDialog(accountManager);
			if (!targetAccountId) {
				logger.debug("Account selection cancelled");
				return;
			}
			const account = accountManager.getAccount(targetAccountId);
			if (account) {
				targetPartition = account.partition;
			}
		}
	}

	const composeWindow = new BrowserWindow({
		width: 1000,
		height: 800,
		backgroundColor: isDarkMode ? "#302a75" : "#fff",
		show: false,
		autoHideMenuBar: true,
		icon: iconChooser.getFile(),
		webPreferences: {
			partition: targetPartition,
			preload: path.join(__dirname, "..", "browser", "index.js"),
			nodeIntegration: false,
			contextIsolation: true,
			sandbox: true,
			spellcheck: false,
			additionalArguments: targetAccountId
				? [`--accountId=${targetAccountId}`]
				: [],
		},
	});

	require("../webauthn").attachWindow(composeWindow.webContents);
	composeWindow.webContents.setWindowOpenHandler(({ url }) => {
		const { shell } = require("electron");
		if (isSafeExternalUrl(url)) shell.openExternal(url);
		return { action: "deny" };
	});
	composeWindow.webContents.on("will-navigate", (event, nextUrl) => {
		if (isExternalNavigation(nextUrl, composeWindow.webContents.getURL(), config)) {
			event.preventDefault();
			if (isSafeExternalUrl(nextUrl)) require("electron").shell.openExternal(nextUrl);
		}
	});
	composeWindow.webContents.on("zoom-changed", (_event, zoomDirection) => {
		composeWindow.webContents.send("zoom-changed", zoomDirection);
	});

	composeWindow.once("ready-to-show", () => {
		composeWindow.show();
	});

	composeWindow.loadURL(url, { userAgent: config.chromeUserAgent });

	logger.debug("Compose window opened with URL:", url);
}

/**
 * Convert mailto: URL to Outlook compose URL
 * @param {string} mailtoUrl - mailto URL (e.g., mailto:user@example.com?subject=Hello)
 * @returns {string} Outlook compose URL
 */
function convertMailtoToOutlookURL(mailtoUrl) {
	try {
		// Remove 'mailto:' prefix
		const mailtoContent = mailtoUrl.substring(7);

		// Parse the mailto URL
		const [recipient, queryString] = mailtoContent.split("?");

		// Build Outlook compose URL
		let outlookUrl = config.url;
		if (!outlookUrl.endsWith("/")) {
			outlookUrl += "/";
		}
		outlookUrl += "mail/deeplink/compose?";

		// Add recipient
		if (recipient) {
			outlookUrl += `to=${encodeURIComponent(recipient)}`;
		}

		// Add other parameters
		// NOTE: Outlook uses ? for cc/bcc and & for subject/body (weird format)
		if (queryString) {
			const params = new URLSearchParams(queryString);

			// cc and bcc use ? separator
			if (params.has("cc")) {
				outlookUrl += `?cc=${params.get("cc")}`;
			}
			if (params.has("bcc")) {
				outlookUrl += `?bcc=${params.get("bcc")}`;
			}

			// subject and body use & separator
			if (params.has("subject")) {
				outlookUrl += `&subject=${encodeURIComponent(params.get("subject"))}`;
			}
			if (params.has("body")) {
				outlookUrl += `&body=${encodeURIComponent(params.get("body"))}`;
			}
		}

		logger.debug("Converted mailto URL to:", outlookUrl);
		return outlookUrl;
	} catch (err) {
		logger.error("Error converting mailto URL:", err);
		return config.url;
	}
}

// NOTE: Old single-window event handlers removed - now handled per-account in AccountManager
// - onBeforeRequestHandler (webRequest handler)
// - onNewWindow (window open handler)
// - onBeforeInput (input event handler)
// - onContextMenu (context menu handler)
// - secureOpenLink (link handling)
// - isOutlookWindow (URL checking)
// - isTelemetryUrl (telemetry blocking)
// - openInBrowser (browser opening)
// - getLinkAction (user dialog for links)
// - removePopupWindowMenu (popup menu removal)
// - createWindow (window creation)
// - createNewBrowserWindow (browser window creation)
