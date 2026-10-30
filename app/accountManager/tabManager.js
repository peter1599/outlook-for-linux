const { isSafeExternalUrl, isExternalNavigation } = require('../security');

/**
 * TabManager - Manages tabs within a single main window
 * Each tab represents one account with isolated session
 */

/**
 * @typedef {Object} Tab
 * @property {string} id - Account ID
 * @property {Electron.WebContents} webContents - The web contents for this tab
 * @property {HTMLElement} element - The tab bar element for this tab
 * @property {Object} account - The account object
 */

class TabManager {
	/**
	 * @param {Electron.BrowserWindow} mainWindow - The main window containing tabs
	 * @param {Object} accountManager - Reference to AccountManager
	 * @param {Object} config - App configuration
	 */
	constructor(mainWindow, accountManager, config) {
		/** @type {Electron.BrowserWindow} */
		this.mainWindow = mainWindow;

		/** @type {Object} */
		this.accountManager = accountManager;

		/** @type {Object} */
		this.config = config;

		/** @type {Map<string, Tab>} */
		this.tabs = new Map();

		/** @type {string|null} */
		this.activeTabId = null;

		/** @type {string|null} */
		this.tabContainerId = "tab-list";

		// Track if tab bar has been injected
		/** @type {boolean} */
		this.tabBarInjected = false;

		console.log("[TabManager] Initialized");
	}

	/**
	 * Create a new tab for an account
	 * @param {Object} account - Account to create tab for
	 * @returns {Electron.WebContents} The created web contents
	 */
	createTab(account) {
		const { BrowserView, session } = require("electron");
		const path = require("path");

		// Set display name based on existing tab count
		const existingCount = this.tabs.size;
		account.displayName = `Account ${existingCount + 1}`;

		console.log(
			`[TabManager] Creating tab for account: ${account.displayName}`,
		);

		// GLM-5 FIX 1: Pre-initialize session partition BEFORE creating BrowserView
		const ses = session.fromPartition(account.partition);
		ses.setUserAgent(this.config.chromeUserAgent);
		this.accountManager.attachIntuneSession(ses);
		console.log(`[TabManager] Session pre-initialized: ${account.partition}`);

		// Create BrowserView for this tab with its own webContents
		const browserView = new BrowserView({
			webPreferences: {
				partition: account.partition,
				preload: path.join(__dirname, "..", "browser", "index.js"),
				additionalArguments: [`--accountId=${account.id}`], // GLM-5 FIX 2: Removed '--tab-mode'
				plugins: true, // CRITICAL: Required for Outlook auth
				nodeIntegration: false,
				contextIsolation: true,
				sandbox: true,
				spellcheck: false,
			},
		});

		const tabWebContents = browserView.webContents;

		// Add BrowserView to main window
		this.mainWindow.addBrowserView(browserView);

		require("../webauthn").attachWindow(tabWebContents);
		tabWebContents.on("zoom-changed", (_event, zoomDirection) => {
			tabWebContents.send("zoom-changed", zoomDirection);
		});

        // Handle opening links in external browser. OAuth popups (e.g. the
        // Mimecast login dialog) call window.open() with explicit popup
        // dimensions and rely on window.opener to postMessage the result back,
        // so ONLY sized popups open as in-app child windows sharing this tab's
        // session; ordinary links (window.open with only 'noopener,noreferrer')
        // still go to the external browser. A Ctrl+Click also requests a sized
        // popup and therefore stays inside the app.
        tabWebContents.setWindowOpenHandler((details) => {
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
		tabWebContents.on('will-navigate', (event, url) => {
			if (isExternalNavigation(url, tabWebContents.getURL(), this.config)) {
				event.preventDefault();
				if (isSafeExternalUrl(url)) require('electron').shell.openExternal(url);
			}
		});

		// Forward hovered link URLs to the renderer for the link preview tooltip
		tabWebContents.on("update-target-url", (_event, url) => {
			tabWebContents.send("hover-link-url", { url: url || "" });
		});

		// Add context menu (right-click) for this tab
		tabWebContents.on("context-menu", (_event, params) => {
			this.showContextMenu(tabWebContents, params);
		});

		// Listen for scale factor changes (when moving to monitors with different DPI)
		tabWebContents.on("preferred-frame-size-changed", () => {
			if (this.activeTabId === account.id) {
				this.updateTabBounds();
			}
		});

		// Store tab reference
		this.tabs.set(account.id, {
			id: account.id,
			webContents: tabWebContents,
			element: browserView,
			account: account,
		});

		// GLM-5 FIX 3: Wait for tab bar to be ready, then use setImmediate for reliable timing
		const waitForTabBar = () => {
			if (this.tabBarInjected) {
				// Switch to tab (makes it visible)
				this.switchTab(account.id);
				// Use setImmediate for more reliable timing than setTimeout
				setImmediate(() => {
					tabWebContents.loadURL(this.config.url);
				});
			} else {
				setTimeout(waitForTabBar, 100);
			}
		};
		waitForTabBar();

		return tabWebContents;
	}

	/**
	 * Switch to a specific tab
	 * @param {string} accountId - The account ID to switch to
	 */
	switchTab(accountId) {
		if (!this.tabs.has(accountId)) {
			console.warn(`[TabManager] Tab not found: ${accountId}`);
			return;
		}

		const tab = this.tabs.get(accountId);
		if (this.activeTabId === accountId) {
			console.log(
				`[TabManager] Tab already active: ${tab.account.displayName}`,
			);
			return;
		}

		// Hide all tabs, show selected one
		this.tabs.forEach((tab, id) => {
			if (id === accountId) {
				// Show this tab
				this.setTabVisibility(tab, true);
			} else {
				// Hide other tabs
				this.setTabVisibility(tab, false);
			}
		});

		this.activeTabId = accountId;
		this.updateTabUI();

		// Show and focus main window
		if (this.mainWindow.isMinimized()) {
			this.mainWindow.restore();
		}
		this.mainWindow.show();
		this.mainWindow.focus();

		console.log(`[TabManager] Switched to tab: ${tab.account.displayName}`);
	}

	/**
	 * Close a tab (tab cleanup only - doesn't remove account)
	 * @param {string} accountId - The account ID to close
	 */
	closeTab(accountId) {
		if (!this.tabs.has(accountId)) {
			console.warn(`[TabManager] Tab not found: ${accountId}`);
			return;
		}

		const tab = this.tabs.get(accountId);
		const wasActive = this.activeTabId === accountId;

		// Remove BrowserView from main window
		if (tab.element) {
			this.mainWindow.removeBrowserView(tab.element);
		}

		// Destroy webContents
		tab.webContents.destroy();

		// Remove from tracking
		this.tabs.delete(accountId);

		// Update tab bar UI to reflect the removal
		this.updateTabUI();

		// If this was active tab, switch to another
		if (wasActive) {
			const remainingIds = Array.from(this.tabs.keys());
			if (remainingIds.length > 0) {
				this.switchTab(remainingIds[0]);
			} else {
				this.activeTabId = null;
			}
		}

		console.log(`[TabManager] Closed tab: ${tab.account.displayName}`);
	}

	/**
	 * Set tab visibility via BrowserView
	 * @param {Tab} tab - The tab to update
	 * @param {boolean} visible - Whether the tab should be visible
	 */
	setTabVisibility(tab, visible) {
		if (!tab.element) return;

		const contentBounds = this.mainWindow.getContentBounds();

		if (visible) {
			// Disable auto-resize (buggy with DPI changes) - we'll update manually
			tab.element.setAutoResize({ width: false, height: false });
			// Set valid bounds to make visible
			this.setBoundsForTab(tab, contentBounds);
		} else {
			// Set bounds outside visible area to hide
			tab.element.setBounds({
				x: -10,
				y: -10,
				width: 0,
				height: 0,
			});
		}
	}

	/**
	 * Set bounds for a visible tab
	 * @param {Tab} tab - The tab to update
	 * @param {Object} bounds - Content bounds {x, y, width, height}
	 */
	setBoundsForTab(tab, bounds) {
		tab.element.setBounds({
			x: 0,
			y: 28,
			width: bounds.width,
			height: bounds.height - 28,
		});
	}

	/**
	 * Update active tab bounds (for DPI scale changes)
	 */
	updateTabBounds() {
		const activeTab = this.getActiveTab();
		if (activeTab && activeTab.element) {
			const contentBounds = this.mainWindow.getContentBounds();
			this.setBoundsForTab(activeTab, contentBounds);
			console.log("[TabManager] Updated tab bounds:", contentBounds);
		}
	}

	/**
	 * Update tab bar UI in the renderer
	 */
	updateTabUI() {
		if (!this.tabBarInjected) {
			return;
		}

		// Send update to renderer - get fresh display name from AccountManager
		const tabsData = Array.from(this.tabs.values()).map((tab) => {
			// Get fresh account data to pick up any display name changes
			const account = this.accountManager.getAccount(tab.id);
			return {
				id: tab.id,
				displayName: account ? account.displayName : tab.account.displayName,
				isActive: tab.id === this.activeTabId,
			};
		});

		this.mainWindow.webContents.send("update-tab-bar", {
			activeTabId: this.activeTabId,
			tabs: tabsData,
		});
	}

	/**
	 * Get currently active tab
	 * @returns {Tab|null}
	 */
	getActiveTab() {
		if (!this.activeTabId) return null;
		return this.tabs.get(this.activeTabId);
	}

	/**
	 * Get all tabs
	 * @returns {Array<Tab>}
	 */
	getAllTabs() {
		return Array.from(this.tabs.values());
	}

	/**
	 * Get a specific tab
	 * @param {string} accountId - The account ID
	 * @returns {Tab|undefined}
	 */
	getTab(accountId) {
		return this.tabs.get(accountId);
	}

	/**
	 * Show context menu for a tab
	 * @param {Electron.WebContents} webContents
	 * @param {Object} params
	 */
	showContextMenu(webContents, params) {
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
					webContents.reload();
				},
			}),
		);

		// Add "Go to Home" option
		menu.append(
			new MenuItem({
				label: "Go to Home",
				accelerator: "Ctrl+Home",
				click: () => {
					webContents.loadURL(this.config.url);
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
							webContents.session.addWordToSpellCheckerDictionary(
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
						webContents.inspectElement(params.x, params.y);
					},
				}),
			);
		}

		menu.popup({ window: this.mainWindow });
	}

	/**
	 * Clean up resources
	 */
	destroy() {
		console.log("[TabManager] Destroying...");

		// Destroy all webContents
		this.tabs.forEach((tab) => {
			tab.webContents.destroy();
		});

		// Clear tracking
		this.tabs.clear();
		this.activeTabId = null;

		console.log("[TabManager] Destroyed");
	}
}

module.exports = TabManager;
