const { BrowserWindow, app, ipcMain, powerMonitor } = require("electron");
const path = require("path");
const windowStateKeeper = require("electron-window-state");
const { LucidLog } = require("lucid-log");
const TrayIconChooser = require("../browser/tools/trayIconChooser");
const TabManager = require("./tabManager");
const { isSafeExternalUrl, isExternalNavigation } = require("../security");

/**
 * Manages multiple Outlook accounts with isolated sessions
 */
class AccountManager {
	/**
	 * @param {Object} config - App configuration
	 * @param {string} iconPath - Path to tray icon
	 * @param {Object} menus - Menus instance
	 * @param {Object} appConfig - App configuration with settings store
	 */
	constructor(config, iconPath, menus, appConfig) {
		/** @type {Object} */
		this.config = config;

		/** @type {string} */
		this.iconPath = iconPath;

		/** @type {Object} */
		this.menus = menus;

		/** @type {Object} */
		this.appConfig = appConfig;
		this.intune = config.auth?.intune?.enabled ? require("../intune") : null;

		/** @type {Array<Account>} */
		this.accounts = [];

		/** @type {boolean} */
		this.isQuitting = false;

		/** @type {LucidLog} */
		this.logger = new LucidLog({
			levels: config.appLogLevels.split(","),
		});

		/** @type {TrayIconChooser} */
		this.iconChooser = new TrayIconChooser(config);

		/** @type {TabManager|null} */
		this.tabManager = null;

		/** @type {BrowserWindow|null} */
		this.mainTabbedWindow = null;

		/** @type {boolean} */
		this.tabbedMode = config.tabbedMode || false;

		// Track when main tabbed window is fully ready (for timing tests)
		/** @type {boolean} */
		this.mainWindowReady = false;

		// Track if app is quitting (to allow main tabbed window to close)
		/** @type {boolean} */
		this.allowQuit = false;

		// Load saved accounts
		this.loadAccounts();

		// Track which partition is currently focused (for badge rotation)
		this.focusedAccountId = null;
		this.badgeRotationInterval = null;
		this.currentBadgeIndex = 0;

		// Initialize tab manager if tabbed mode is enabled
		if (this.tabbedMode) {
			this.initializeTabbedMode();
		}

		// Set this AccountManager on the Menus and Tray
		if (menus) {
			menus.setAccountManager(this);
			if (menus.tray) {
				menus.tray.setAccountManager(this);
			}
		}

		// Listen for before-quit to allow windows to close properly
		app.on("before-quit", () => {
			this.isQuitting = true;
		});

		// Register global handlers once (shared by all windows)
		this.registerGlobalHandlers();
	}

	attachIntuneSession(browserSession) {
		if (this.intune) {
			this.intune.attachToSession(browserSession);
		}
	}

	/**
	 * Register global event handlers (only once for all accounts)
	 */
	registerGlobalHandlers() {
		// Check if already registered
		if (this.globalHandlersRegistered) {
			return;
		}
		this.globalHandlersRegistered = true;

		const refreshAllWindows = () => {
			this.accounts.forEach((account) => {
				if (account.window && !account.window.isDestroyed()) {
					this.refreshAccountWindow(account);
				}
			});
		};

		// Register global offline-retry handler
		ipcMain.removeAllListeners("offline-retry");
		ipcMain.on("offline-retry", refreshAllWindows);

		// Register global system resume handler
		powerMonitor.removeAllListeners("resume");
		powerMonitor.on("resume", refreshAllWindows);

		this.logger.info("Global handlers registered for connection management");
	}

	/**
	 * Refresh a single account window
	 * @param {Account} account
	 */
	refreshAccountWindow(account) {
		// Skip in tabbed mode - tabs handle their own refreshing
		if (this.tabbedMode) {
			this.logger.info("Tabbed mode: skipping window refresh");
			return;
		}

		if (!account.window || account.window.isDestroyed()) {
			return;
		}

		const currentUrl = account.window.webContents.getURL();
		const hasUrl =
			currentUrl && currentUrl.startsWith("https://") ? true : false;

		if (hasUrl) {
			account.window.reload();
		} else {
			account.window.loadURL(this.config.url, {
				userAgent: this.config.chromeUserAgent,
			});
		}
	}

	/**
	 * Load accounts from persistent storage
	 */
	loadAccounts() {
		const savedAccounts = this.appConfig.settingsStore.get("accounts", []);
		this.accounts = savedAccounts.map((acc) => ({
			...acc,
			window: null,
			unreadCount: 0,
			reminderCount: 0,
		}));
		this.logger.info(`Loaded ${this.accounts.length} accounts from storage`);
	}

	/**
	 * Save accounts to persistent storage
	 */
	saveAccounts() {
		const accountsToSave = this.accounts.map((acc) => ({
			id: acc.id,
			email: acc.email,
			partition: acc.partition,
			displayName: acc.displayName,
			autoRestore: acc.autoRestore !== false, // default true
			createdAt: acc.createdAt,
		}));
		this.appConfig.settingsStore.set("accounts", accountsToSave);
		this.logger.debug(`Saved ${accountsToSave.length} accounts to storage`);
	}

	/**
	 * Generate unique account ID
	 * @returns {string}
	 */
	generateAccountId() {
		return `account-${Date.now()}-${Math.random().toString(36).substr(2, 9)}`;
	}

	/**
	 * Create a new account
	 * @param {Object} options
	 * @param {string} [options.email] - Optional email address
	 * @param {string} [options.displayName] - Optional display name
	 * @returns {Account} The created account
	 */
	initializeTabbedMode() {
		// Prevent double initialization
		if (this.mainTabbedWindow && !this.mainTabbedWindow.isDestroyed()) {
			this.logger.info("Tabbed mode already initialized, skipping");
			return;
		}

		this.logger.info("Initializing tabbed mode");

		const { BrowserWindow, nativeTheme } = require("electron");
		const path = require("path");

		// Create main tabbed window
		const mainWindow = new BrowserWindow({
			width: 1400,
			height: 900,
			x11Class: "outlook-for-linux",
			backgroundColor: nativeTheme.shouldUseDarkColors ? "#302a75" : "#ffffff",
			show: false,
			autoHideMenuBar: this.config.menubar === "auto",
			icon: this.iconChooser.getFile(),
			webPreferences: {
				partition: "persist:outlook-tabbed-ui",
				nodeIntegration: false,
				contextIsolation: true,
				sandbox: true,
				preload: path.join(__dirname, "tabbar-preload.js"),
			},
		});

		// Store reference
		this.mainTabbedWindow = mainWindow;

		// Handle window state BEFORE creating TabManager
		const windowStateKeeper = require("electron-window-state");
		const mainWindowState = windowStateKeeper({
			defaultWidth: 1200,
			defaultHeight: 800,
		});

		// Update window bounds from saved state
		mainWindow.setBounds({
			width: mainWindowState.width,
			height: mainWindowState.height,
			x: mainWindowState.x,
			y: mainWindowState.y,
		});

		// Create TabManager
		this.tabManager = new TabManager(mainWindow, this, this.config);

		// Let window state keeper manage the window
		mainWindowState.manage(mainWindow);

		// Prevent eval
		mainWindow.eval = global.eval = function () {
			// eslint-disable-line no-eval
			throw new Error("Sorry, this app does not support window.eval().");
		};

		// Load tab bar UI into main window
		const tabBarPath = path.join(__dirname, "tabbar.html");
		mainWindow.loadFile(tabBarPath);

		// Mark tab bar as injected when loaded
		mainWindow.webContents.on("did-finish-load", () => {
			if (!this.tabManager.tabBarInjected) {
				this.tabManager.tabBarInjected = true;
				this.logger.info("Tab bar loaded and injected");
				// Update tab UI with initial state (empty initially)
				this.tabManager.updateTabUI();
			}
		});

		// Show when ready
		mainWindow.once("ready-to-show", () => {
			this.mainWindowReady = true;
			this.logger.info("Main tabbed window ready-to-show fired");
			mainWindow.show();
		});

		// Handle window close button - respect closeAppOnCross setting and allow quit
		mainWindow.on("close", (event) => {
			if (!this.config.closeAppOnCross && !this.isQuitting) {
				event.preventDefault();
				mainWindow.hide();
			}
		});

		// Handle DPI scale changes when moving between monitors
		// 'moved' fires only after drag completes, not continuously during drag
		mainWindow.on("moved", () => {
			if (this.tabManager) {
				this.tabManager.updateTabBounds();
			}
		});

		// Debounced resize handler (resize fires continuously during drag)
		let resizeTimeout;
		mainWindow.on("resize", () => {
			if (this.tabManager) {
				clearTimeout(resizeTimeout);
				resizeTimeout = setTimeout(() => {
					this.tabManager.updateTabBounds();
				}, 100);
			}
		});

		// Handle window closed
		mainWindow.on("closed", () => {
			this.logger.info("Main tabbed window closed");
			this.mainTabbedWindow = null;
			if (this.tabManager) {
				this.tabManager.destroy();
			}
		});

		this.logger.info("Tabbed mode initialized");
	}

	/**
	 * Create a new account
	 * @param {Object} options
	 * @param {string} [options.email] - Optional email address
	 * @param {string} [options.displayName] - Optional display name
	 * @returns {Account} The created account
	 */
	createAccount(options = {}) {
		// If in tabbed mode, use tab manager instead
		if (this.tabbedMode && this.tabManager) {
			this.logger.info(
				`Creating tab for account: ${options.displayName || "New Account"}`,
			);
			const id = this.generateAccountId();
			const partition = `persist:${id}`;

			// Create account object and save to accounts array
			const account = {
				id,
				partition,
				email: options.email || null,
				displayName:
					options.displayName || `Account ${this.accounts.length + 1}`,
				autoRestore: true,
				createdAt: Date.now(),
				window: null, // No window in tabbed mode
				unreadCount: 0,
				reminderCount: 0,
			};

			this.accounts.push(account);
			this.saveAccounts();

			// Create the tab
			this.tabManager.createTab(account);

			this.updateTrayMenu();
			return account;
		}

		// Otherwise, use existing window-based approach
		const id = this.generateAccountId();
		const partition = `persist:${id}`;

		const account = {
			id,
			partition,
			email: options.email || null,
			displayName: options.displayName || `Account ${this.accounts.length + 1}`,
			autoRestore: true,
			createdAt: Date.now(),
			window: null,
			unreadCount: 0,
			reminderCount: 0,
		};

		this.accounts.push(account);
		this.saveAccounts();

		this.createAccountWindow(account);

		this.logger.info(`Created account: ${account.displayName} (${account.id})`);
		this.updateTrayMenu();

		return account;
	}

	/**
	 * Create the browser window for an account
	 * @param {Account} account
	 */
	createAccountWindow(account) {
		// In tabbed mode, delegate to TabManager
		if (this.tabbedMode && this.tabManager) {
			this.logger.info(`Creating tab for account: ${account.displayName}`);
			// Check if tab already exists
			const existingTab = this.tabManager.getTab(account.id);
			if (existingTab) {
				this.tabManager.switchTab(account.id);
				return;
			}
			// Create new tab for this account
			this.tabManager.createTab({
				id: account.id,
				displayName: account.displayName,
				email: account.email,
				partition: account.partition,
			});
			return;
		}

		// Check if window already exists
		if (account.window && !account.window.isDestroyed()) {
			account.window.show();
			account.window.focus();
			return;
		}

		// Load window state
		const windowState = windowStateKeeper({
			defaultWidth: 1280,
			defaultHeight: 800,
			file: `window-state-${account.id}`,
		});

		const window = new BrowserWindow({
			title: `Microsoft Outlook - ${account.displayName}`,
			x11Class: "outlook-for-linux",
			x: windowState.x,
			y: windowState.y,
			width: windowState.width,
			height: windowState.height,
			backgroundColor: "#302a75",
			show: false,
			autoHideMenuBar: this.config.menubar == "auto",
			icon: this.iconChooser.getFile(),
			webPreferences: {
				partition: account.partition,
				preload: path.join(__dirname, "..", "browser", "index.js"),
				plugins: true,
				nodeIntegration: false,
				contextIsolation: true,
				sandbox: true,
				spellcheck: false,
				additionalArguments: [`--accountId=${account.id}`],
			},
		});

		require("../webauthn").attachWindow(window.webContents);
		window.webContents.on("zoom-changed", (_event, zoomDirection) => {
			window.webContents.send("zoom-changed", zoomDirection);
		});
		this.attachIntuneSession(window.webContents.session);
        // Add-in / OAuth popups (e.g. the Mimecast login dialog) authenticate by
        // calling window.open() with explicit popup dimensions and postMessage
        // the result back to window.opener. ONLY those sized popups should open
        // inside the app; ordinary links (including window.open(url, '_blank',
        // 'noopener,noreferrer')) carry no width/height and must go to the
        // external browser. A Ctrl+Click from the preload also requests a sized
        // popup, so it too stays inside the app.
        window.webContents.setWindowOpenHandler((details) => {
			const { shell } = require('electron');
            const { url, features } = details;
            const featureStr = typeof features === 'string' ? features : '';
            const isAuthPopup =
                /(?:^|[,\s])(?:width|height|innerwidth|innerheight)\s*=/i.test(
                    featureStr,
                );
            if (isAuthPopup && (isSafeExternalUrl(url) || url === 'about:blank')) {
                return {
                    action: 'allow',
                    overrideBrowserWindowOptions: {
                        autoHideMenuBar: true,
                        webPreferences: {
                            partition: account.partition,
                            preload: path.join(__dirname, '..', 'browser', 'index.js'),
                            nodeIntegration: false,
                            contextIsolation: true,
                            sandbox: true,
                        },
                    },
                };
            }
			if (isSafeExternalUrl(url)) shell.openExternal(url);
			return { action: 'deny' };
		});
		window.webContents.on('will-navigate', (event, url) => {
			if (isExternalNavigation(url, window.webContents.getURL(), this.config)) {
				event.preventDefault();
				if (isSafeExternalUrl(url)) require('electron').shell.openExternal(url);
			}
		});

		// Store reference
		account.window = window;

		// Manage window state
		windowState.manage(window);

		// Prevent eval
		window.eval = global.eval = function () {
			// eslint-disable-line no-eval
			throw new Error("Sorry, this app does not support window.eval().");
		};

		// Show when ready, or after a timeout fallback
		let shown = false;
		const showWindow = () => {
			if (!shown) {
				shown = true;
				window.show();
				this.logger.info(`Account window shown: ${account.displayName}`);
			}
		};

		window.once("ready-to-show", showWindow);

		// Fallback: show window after 5 seconds even if not ready
		setTimeout(() => {
			showWindow();
		}, 5000);

		// Handle window close - prevent closing and hide instead (minimize to tray)
		// But allow closing if app is quitting
		window.on("close", (event) => {
			if (this.isQuitting) {
				// Allow the window to close when app is quitting
				this.logger.info(
					`Account window closing (app quitting): ${account.displayName}`,
				);
				return;
			}

			this.logger.info(
				`Account window close requested: ${account.displayName}`,
			);
			// Prevent the window from closing, just hide it
			event.preventDefault();
			window.hide();
			this.logger.info(`Account window hidden to tray: ${account.displayName}`);
		});

		// Handle window closed
		window.on("closed", () => {
			account.window = null;
			this.logger.info(`Account window closed: ${account.displayName}`);
		});

		// Handle window focus for badge rotation
		window.on("focus", () => {
			this.focusedAccountId = account.id;
		});

		// Forward hovered link URLs to the renderer for the link preview tooltip
		window.webContents.on("update-target-url", (_event, url) => {
			window.webContents.send("hover-link-url", { url: url || "" });
		});

		// Add context menu (right-click)
		window.webContents.on("context-menu", (_event, params) => {
			const { Menu, MenuItem, clipboard, shell } = require("electron");
			const menu = new Menu();

			// Link-specific options
			if (params.linkURL) {
				menu.append(
					new MenuItem({
						label: "Open Link in Browser",
						click: () => shell.openExternal(params.linkURL),
					}),
				);
				menu.append(
					new MenuItem({
						label: "Copy Link",
						click: () => clipboard.writeText(params.linkURL),
					}),
				);
				menu.append(new MenuItem({ type: "separator" }));
			}

			// Add "Reload Page" option
			menu.append(
				new MenuItem({
					label: "Reload Page",
					accelerator: "Ctrl+R",
					click: () => {
						window.reload();
					},
				}),
			);

			// Add "Go to Home" option
			menu.append(
				new MenuItem({
					label: "Go to Home",
					accelerator: "Ctrl+Home",
					click: () => {
						window.loadURL(this.config.url, {
							userAgent: this.config.chromeUserAgent,
						});
					},
				}),
			);

			menu.append(new MenuItem({ type: "separator" }));

			// Standard context menu items (if text is selected or in an input field)
			if (params.isEditable || params.selectionText) {
				if (params.misspelledWord) {
					menu.append(
						new MenuItem({
							label: "Add to Dictionary",
							click: () => {
								window.webContents.session.addWordToSpellCheckerDictionary(
									params.misspelledWord,
								);
							},
						}),
					);
					menu.append(new MenuItem({ type: "separator" }));
				}

				if (params.isEditable) {
					menu.append(new MenuItem({ label: "Cut", role: "cut" }));
					menu.append(new MenuItem({ label: "Copy", role: "copy" }));
					menu.append(new MenuItem({ label: "Paste", role: "paste" }));
				} else if (params.selectionText) {
					menu.append(new MenuItem({ label: "Copy", role: "copy" }));
				}

				menu.append(new MenuItem({ type: "separator" }));
			}

			// Add "Inspect Element" for debugging
			if (this.config.webDebug) {
				menu.append(
					new MenuItem({
						label: "Inspect Element",
						click: () => {
							window.webContents.inspectElement(params.x, params.y);
						},
					}),
				);
			}

			menu.popup({ window });
		});

		// Load the Outlook URL for this account window
		window.loadURL(this.config.url, { userAgent: this.config.chromeUserAgent });

		this.logger.info(
			`Created account window: ${account.displayName} (partition: ${account.partition})`,
		);
	}

	/**
	 * Remove an account
	 * @param {string} accountId
	 */
	removeAccount(accountId) {
		const index = this.accounts.findIndex((a) => a.id === accountId);
		if (index === -1) {
			this.logger.warn(`Account not found: ${accountId}`);
			return;
		}

		const account = this.accounts[index];

		// Close window if open
		if (account.window && !account.window.isDestroyed()) {
			account.window.close();
		}

		// In tabbed mode, close the tab
		if (this.tabbedMode && this.tabManager) {
			this.tabManager.closeTab(accountId);
		}

		// Remove from list
		this.accounts.splice(index, 1);
		this.saveAccounts();

		this.logger.info(`Removed account: ${account.displayName}`);
		this.updateTrayMenu();

		// If no accounts left, create a new one
		if (this.accounts.length === 0) {
			this.logger.info("No accounts remaining, creating default account");
			this.createAccount();
		}
	}

	/**
	 * Get account by ID
	 * @param {string} accountId
	 * @returns {Account|null}
	 */
	getAccount(accountId) {
		return this.accounts.find((a) => a.id === accountId) || null;
	}

	/**
	 * Get account by window
	 * @param {BrowserWindow} window
	 * @returns {Account|null}
	 */
	getAccountByWindow(window) {
		return this.accounts.find((a) => a.window === window) || null;
	}

	/**
	 * Get all accounts
	 * @returns {Array<Account>}
	 */
	getAllAccounts() {
		return [...this.accounts];
	}

	/**
	 * Restore accounts (open windows for accounts with autoRestore enabled)
	 */
	async restoreAccounts() {
		const accountsToRestore = this.accounts.filter(
			(a) => a.autoRestore !== false,
		);

		if (accountsToRestore.length === 0 && this.accounts.length === 0) {
			// No accounts exist, create first one
			this.logger.info("No accounts found, creating default account");
			// In tabbed mode, wait for main window to be ready before creating first account
			if (this.tabbedMode && this.mainTabbedWindow) {
				this.logger.info(
					"Tabbed mode: waiting for main window ready-to-show before creating first account",
				);
				this.mainTabbedWindow.once("ready-to-show", () => {
					this.logger.info("Main window ready, creating first account");
					this.createAccount();
				});
				// Show the window to trigger ready-to-show
				this.mainTabbedWindow.show();
			} else {
				this.createAccount();
			}
			return;
		}

		this.logger.info(`Restoring ${accountsToRestore.length} accounts...`);

		// TIMING TEST: In tabbed mode, wait for main window to be ready before restoring tabs
		if (this.tabbedMode && !this.mainWindowReady) {
			this.logger.info(
				"Tabbed mode: waiting for main window ready-to-show before restoring accounts",
			);
			this.mainTabbedWindow.once("ready-to-show", () => {
				this.logger.info("Main window ready, now restoring accounts");
				for (const account of accountsToRestore) {
					this.createAccountWindow(account);
				}
				this.startBadgeRotation();
			});
		} else {
			for (const account of accountsToRestore) {
				this.createAccountWindow(account);
			}
			// Start badge rotation
			this.startBadgeRotation();
		}
	}

	/**
	 * Update account email (detected from page)
	 * @param {string} accountId
	 * @param {string} email
	 */
	setAccountEmail(accountId, email) {
		const account = this.getAccount(accountId);
		if (account && account.email !== email) {
			account.email = email;
			this.saveAccounts();
			this.logger.info(
				`Updated account email: ${account.displayName} -> ${email}`,
			);
			this.updateTrayMenu();
		}
	}

	/**
	 * Update account display name
	 * @param {string} accountId
	 * @param {string} displayName
	 */
	setAccountDisplayName(accountId, displayName) {
		const account = this.getAccount(accountId);
		if (account) {
			// If displayName is empty or whitespace, clear manual override to allow detection
			if (displayName.trim() === "") {
				account.manualDisplayName = null;
				// Revert to detected email or default
				account.displayName =
					account.email || `Account ${this.accounts.indexOf(account) + 1}`;
			} else {
				// Set manual override - detection won't override this
				account.manualDisplayName = displayName;
				account.displayName = displayName;
			}
			this.saveAccounts();

			// Update window title
			if (account.window && !account.window.isDestroyed()) {
				account.window.setTitle(`Microsoft Outlook - ${account.displayName}`);
			}

			// Update tab UI in tabbed mode
			if (this.tabbedMode && this.tabManager) {
				this.tabManager.updateTabUI();
			}

			this.updateTrayMenu();
		}
	}

	/**
	 * Toggle auto-restore setting for an account
	 * @param {string} accountId
	 */
	toggleAutoRestore(accountId) {
		const account = this.getAccount(accountId);
		if (account) {
			account.autoRestore = !account.autoRestore;
			this.saveAccounts();
			this.logger.info(
				`Auto-restore ${account.autoRestore ? "enabled" : "disabled"} for: ${account.displayName}`,
			);
			this.updateTrayMenu();
		}
	}

	/**
	 * Focus an account's window or tab
	 * @param {string} accountId
	 */
	focusAccount(accountId) {
		const account = this.getAccount(accountId);
		if (!account) {
			this.logger.warn(`Account not found: ${accountId}`);
			return;
		}

		// In tabbed mode, switch to tab instead
		if (this.tabbedMode && this.tabManager) {
			this.logger.info(`Tabbed mode: switching to tab: ${account.displayName}`);
			this.tabManager.switchTab(accountId);
			this.focusedAccountId = accountId;
			return;
		}

		// Original window-based behavior
		if (account.window && !account.window.isDestroyed()) {
			account.window.show();
			account.window.focus();
		}

		this.focusedAccountId = accountId;
	}

	/**
	 * Show all account windows (cascade them)
	 */
	showAllWindows() {
		// In tabbed mode, just show main tabbed window
		if (
			this.tabbedMode &&
			this.mainTabbedWindow &&
			!this.mainTabbedWindow.isDestroyed()
		) {
			if (this.mainTabbedWindow.isMinimized()) {
				this.mainTabbedWindow.restore();
			}
			this.mainTabbedWindow.show();
			this.mainTabbedWindow.focus();
			return;
		}

		// Original window-based behavior
		let offset = 0;
		for (const account of this.accounts) {
			if (account.window && !account.window.isDestroyed()) {
				account.window.show();
				if (offset > 0) {
					const [x, y] = account.window.getPosition();
					account.window.setPosition(x + offset, y + offset);
				}
				offset += 30;
			}
		}

		// Focus the first account
		if (this.accounts.length > 0) {
			this.focusAccount(this.accounts[0].id);
		}
	}

	/**
	 * Update unread count for an account
	 * @param {string} accountId
	 * @param {number} count
	 */
	updateUnreadCount(accountId, count) {
		const account = this.getAccount(accountId);
		if (account) {
			account.unreadCount = count;
			this.updateBadge();
		}
	}

	/**
	 * Update reminder count for an account
	 * @param {string} accountId
	 * @param {number} count
	 */
	updateReminderCount(accountId, count) {
		const account = this.getAccount(accountId);
		if (account) {
			account.reminderCount = count;
			this.updateBadge();
		}
	}

	/**
	 * Start badge rotation between accounts
	 */
	startBadgeRotation() {
		if (this.badgeRotationInterval) {
			clearInterval(this.badgeRotationInterval);
		}

		this.badgeRotationInterval = setInterval(() => {
			this.updateBadge();
		}, 3000);
	}

	/**
	 * Update badge (rotate between accounts)
	 */
	updateBadge() {
		// Find accounts with unread items
		const accountsWithUnread = this.accounts.filter(
			(a) => a.unreadCount > 0 || a.reminderCount > 0,
		);

		if (accountsWithUnread.length === 0) {
			app.setBadgeCount(0);
			if (this.menus && this.menus.tray) {
				this.menus.tray.updateBadge(0, "email");
			}
			return;
		}

		// Rotate to next account
		this.currentBadgeIndex =
			(this.currentBadgeIndex + 1) % accountsWithUnread.length;
		const account = accountsWithUnread[this.currentBadgeIndex];

		// Show reminder count if available, otherwise email count
		const count =
			account.reminderCount > 0 ? account.reminderCount : account.unreadCount;
		const type = account.reminderCount > 0 ? "reminder" : "email";

		app.setBadgeCount(count);
		if (this.menus && this.menus.tray) {
			this.menus.tray.updateBadge(count, type);
		}
	}

	/**
	 * Update tray menu with account list
	 */
	updateTrayMenu() {
		if (this.menus && this.menus.updateTrayMenu) {
			this.menus.updateTrayMenu();
		}
	}

	/**
	 * Close all account windows
	 */
	closeAllWindows() {
		for (const account of this.accounts) {
			if (account.window && !account.window.isDestroyed()) {
				account.window.close();
			}
		}
	}
}

module.exports = AccountManager;
