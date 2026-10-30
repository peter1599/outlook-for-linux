console.log("[Preload] Script starting...");

const webauthnElectron = require("electron");

/**
 * Runs in the page's main world (webFrame.executeJavaScript) so it replaces
 * navigator.credentials. The sandboxed preload cannot patch the page's objects
 * directly. Requests go to the main process through window.__outlookWebauthn,
 * which exposes only the create and get channels.
 */
function webauthnOverride() {
	const bridge = window.__outlookWebauthn;
	if (!bridge || !navigator.credentials?.create || !navigator.credentials?.get) return;

	const toBase64url = (buffer) => {
		if (!buffer) return "";
		const bytes = buffer instanceof ArrayBuffer ? new Uint8Array(buffer) : buffer;
		let binary = "";
		for (let i = 0; i < bytes.length; i += 8192) {
			binary += String.fromCodePoint(...bytes.subarray(i, i + 8192));
		}
		return btoa(binary).replaceAll("+", "-").replaceAll("/", "_").replace(/={1,2}$/, "");
	};
	const fromBase64url = (value) => {
		let base64 = value.replaceAll("-", "+").replaceAll("_", "/");
		while (base64.length % 4 !== 0) base64 += "=";
		return Uint8Array.from(atob(base64), (c) => c.codePointAt(0)).buffer;
	};
	const timeoutSec = (ms) => (ms ? Math.floor(ms / 1000) : 60);
	const describe = (list) =>
		(list || []).map((c) => ({ id: toBase64url(c.id), type: c.type, transports: c.transports }));
	const withPrototype = (prototype, properties) =>
		Object.create(prototype, Object.getOwnPropertyDescriptors(properties));
	const failure = (message) => {
		if (/invalid/i.test(message) && !/notallowed/i.test(message)) {
			return new DOMException(message, "InvalidStateError");
		}
		return new DOMException(message, "NotAllowedError");
	};

	const createOptions = (publicKey) => ({
		challenge: toBase64url(publicKey.challenge),
		rpId: publicKey.rp?.id || "",
		rpName: publicKey.rp?.name || "",
		userId: toBase64url(publicKey.user?.id),
		userName: publicKey.user?.name || "",
		userDisplayName: publicKey.user?.displayName || "",
		pubKeyCredParams: publicKey.pubKeyCredParams,
		timeout: timeoutSec(publicKey.timeout),
		authenticatorSelection: publicKey.authenticatorSelection || {},
		attestation: publicKey.attestation || "none",
		excludeCredentials: describe(publicKey.excludeCredentials),
	});
	const getOptions = (publicKey) => ({
		challenge: toBase64url(publicKey.challenge),
		rpId: publicKey.rpId || "",
		timeout: timeoutSec(publicKey.timeout),
		userVerification: publicKey.userVerification || "preferred",
		allowCredentials: describe(publicKey.allowCredentials),
	});

	const buildCreateResponse = (data) =>
		withPrototype(PublicKeyCredential.prototype, {
			id: data.credentialId,
			rawId: fromBase64url(data.rawId),
			type: data.type,
			authenticatorAttachment: "cross-platform",
			response: withPrototype(AuthenticatorAttestationResponse.prototype, {
				attestationObject: fromBase64url(data.attestationObject),
				clientDataJSON: fromBase64url(data.clientDataJson),
				getAuthenticatorData: () => fromBase64url(data.authenticatorData),
				getTransports: () => data.transports || ["usb"],
				getPublicKey: () => null,
				getPublicKeyAlgorithm: () => data.publicKeyAlgorithm || -7,
			}),
			getClientExtensionResults: () => ({}),
			toJSON: () => ({
				id: data.credentialId,
				rawId: data.rawId,
				type: data.type,
				response: {
					attestationObject: data.attestationObject,
					clientDataJSON: data.clientDataJson,
				},
			}),
		});

	const buildGetResponse = (data) => {
		const authenticatorData = fromBase64url(data.authenticatorData);
		const response = withPrototype(AuthenticatorAssertionResponse.prototype, {
			authenticatorData,
			clientDataJSON: fromBase64url(data.clientDataJson),
			signature: fromBase64url(data.signature),
			userHandle: data.userHandle ? fromBase64url(data.userHandle) : null,
			getAuthenticatorData: () => authenticatorData,
		});
		return withPrototype(PublicKeyCredential.prototype, {
			id: data.credentialId,
			rawId: fromBase64url(data.rawId),
			type: data.type,
			authenticatorAttachment: "cross-platform",
			response,
			getClientExtensionResults: () => ({}),
			toJSON: () => ({
				id: data.credentialId,
				rawId: data.rawId,
				type: data.type,
				authenticatorAttachment: "cross-platform",
				clientExtensionResults: {},
				response: {
					authenticatorData: data.authenticatorData,
					clientDataJSON: data.clientDataJson,
					signature: data.signature,
					userHandle: data.userHandle || null,
				},
			}),
		});
	};

	const originalCreate = navigator.credentials.create.bind(navigator.credentials);
	const originalGet = navigator.credentials.get.bind(navigator.credentials);

	const relay = async (call, build) => {
		let result;
		try {
			result = await call();
		} catch (err) {
			throw failure(err.message);
		}
		if (!result.success) throw failure(result.error);
		return build(result.data);
	};

	navigator.credentials.create = (options) => {
		if (!options?.publicKey) return originalCreate(options);
		return relay(() => bridge.create(createOptions(options.publicKey)), buildCreateResponse);
	};
	navigator.credentials.get = (options) => {
		if (!options?.publicKey || options.mediation === "conditional") return originalGet(options);
		return relay(() => bridge.get(getOptions(options.publicKey)), buildGetResponse);
	};

	window.addEventListener("message", async (event) => {
		const message = event.data;
		if (message?.type !== "webauthn-request") return;
		const { id, channel, data } = message;
		if (channel !== "webauthn:create" && channel !== "webauthn:get") return;
		const payload = { ...data, frameOrigin: event.origin };
		try {
			const result = channel === "webauthn:create"
				? await bridge.create(payload)
				: await bridge.get(payload);
			event.source.postMessage(
				result.success
					? { type: "webauthn-response", id, result: result.data }
					: { type: "webauthn-response", id, error: result.error },
				event.origin,
			);
		} catch (err) {
			event.source.postMessage({ type: "webauthn-response", id, error: err.message }, event.origin);
		}
	});
}

if (process.platform === "linux") {
	const webauthnSettings = webauthnElectron.ipcRenderer.sendSync("getWebauthnConfigSync");
	if (webauthnSettings?.enabled) {
		webauthnElectron.contextBridge.exposeInMainWorld("__outlookWebauthn", {
			create: (options) => webauthnElectron.ipcRenderer.invoke("webauthn:create", options),
			get: (options) => webauthnElectron.ipcRenderer.invoke("webauthn:get", options),
		});
		webauthnElectron.webFrame.executeJavaScript(`(${webauthnOverride})()`);
	}
}

(async function () {
    const { ipcRenderer, webFrame } = require("electron");

	console.log("[Preload] Inside IIFE, getting config...");

	// Get account ID from additional arguments
	const accountIdArg = process.argv.find((arg) =>
		arg.startsWith("--accountId="),
	);
	const accountId = accountIdArg ? accountIdArg.split("=")[1] : null;
	console.log("[Preload] Account ID:", accountId);

	let config;
	ipcRenderer
		.invoke("getConfig")
		.then((mainConfig) => {
			config = mainConfig;
			console.log(
				"[Preload] Config loaded, disableNotifications:",
				config.disableNotifications,
			);

			// Start account email detection if we have an account ID
			if (accountId) {
				detectAccountEmail(accountId, ipcRenderer);
			}

			// Initialize notification interception FIRST (before other modules that might fail)
			if (!config.disableNotifications) {
				console.log("[Preload] Calling initNotificationInterception...");
				initNotificationInterception(ipcRenderer);
				console.log("[Preload] initNotificationInterception called");

				// Setup unread count observer
				if (accountId) {
					console.log("[Preload] Setting up unread count observer...");
					setupUnreadCountObserver(ipcRenderer, accountId);

					// Setup reminder count observer
					console.log("[Preload] Setting up reminder count observer...");
					setupReminderCountObserver(ipcRenderer, accountId);
				} else {
					// Fallback for single account mode (no accountId)
					console.log(
						"[Preload] Setting up unread count observer (no accountId)...",
					);
					setupUnreadCountObserver(ipcRenderer, null);

					// Setup reminder count observer
					console.log(
						"[Preload] Setting up reminder count observer (no accountId)...",
					);
					setupReminderCountObserver(ipcRenderer, null);
				}
			}

			// Initialize other modules (wrapped in try-catch as they have issues)
			try {
				initializeModules(config, ipcRenderer);
				console.log("[Preload] Modules initialized");
			} catch (err) {
				console.error("[Preload] Error initializing modules:", err);
			}
		})
		.catch((err) => {
			console.error("[Preload] Config error:", err);
		});

	Object.defineProperty(navigator.serviceWorker, "register", {
		value: () => {
			return Promise.reject();
		},
	});

    // Default: links open in the external browser (handled by the main
    // process' window-open handler). Ctrl+Click keeps the link INSIDE the app
    // by requesting a sized popup, which the main process routes to an internal
    // BrowserWindow that shares this account's session.
	document.addEventListener(
		"click",
		(event) => {
			if (!event.ctrlKey) return;

			const anchor = event.target.closest("a[href]");
			if (!anchor) return;

			const url = anchor.href;
			if (!url || !url.startsWith("http")) return;

			event.preventDefault();
			event.stopPropagation();

            // The width/height features mark this as an in-app popup so the
            // main-process handler opens it internally instead of externally.
            window.open(url, "_blank", "width=1200,height=800");
		},
		true,
	);

	// Keep the CustomNotification class for sound playback
	class CustomNotification {
		constructor(title, options) {
			if (config.disableNotifications) {
				return;
			}
			options = options || {};

			const notifSound = {
				type: options.type ? options.type : "new-message",
				audio: "default",
				title: title,
				body: options.body,
			};
			ipcRenderer.invoke("play-notification-sound", notifSound);

			// Note: Native notifications are handled by MutationObserver below
		}

		static requestPermission(callback) {
			if (typeof callback == "function") {
				callback("granted");
			}
		}

		static get permission() {
			return "granted";
		}
	}

	window.Notification = CustomNotification;
})();

// ------------------------------------------------------------
// Link preview (hover tooltip / status bar)
// ------------------------------------------------------------
(function initLinkPreview() {
	const { ipcRenderer, contextBridge } = require("electron");

	contextBridge.exposeInMainWorld("linkPreview", {
		onHover: (callback) => {
			ipcRenderer.on("hover-link-url", (_, payload) => {
				callback(payload);
			});
		},
	});

	const tooltipCss = `
		#link-preview-tooltip {
			position: fixed;
			bottom: 8px;
			left: 8px;
			right: 8px;
			padding: 6px 10px;
			font-size: 12px;
			background: rgba(30, 30, 30, 0.95);
			color: #ddd;
			border-radius: 4px;
			white-space: nowrap;
			overflow: hidden;
			text-overflow: ellipsis;
			z-index: 999999;
			display: none;
		}
	`;

	function createTooltip() {
		const style = document.createElement("style");
		style.textContent = tooltipCss;
		document.head.appendChild(style);

		const tooltip = document.createElement("div");
		tooltip.id = "link-preview-tooltip";
		document.body.appendChild(tooltip);

		window.linkPreview.onHover(({ url }) => {
			if (!url) {
				tooltip.style.display = "none";
				return;
			}
			tooltip.textContent = url;
			tooltip.style.display = "flex";
		});
	}

	if (document.body) {
		createTooltip();
	} else {
		document.addEventListener("DOMContentLoaded", createTooltip);
	}
})();

/**
 * Initialize MutationObserver to intercept Outlook notification elements
 * @param {Electron.IpcRenderer} ipcRenderer
 */
function initNotificationInterception(ipcRenderer) {
	console.log("[Notification] Initializing notification interception...");

	// Wait for DOM to be ready, then setup observer
	const setupObserver = () => {
		// Look for the notification pane container
		const notificationPane = document.querySelector(
			'[data-app-section="NotificationPane"]',
		);

		if (notificationPane) {
			console.log("[Notification] Found NotificationPane, setting up observer");
			observeNotificationPane(notificationPane, ipcRenderer);
		} else {
			// If not found, observe body and wait for it to appear
			console.log(
				"[Notification] NotificationPane not found, observing body for it",
			);
			observeForNotificationPane(ipcRenderer);
		}
	};

	if (document.readyState === "loading") {
		document.addEventListener("DOMContentLoaded", () =>
			setTimeout(setupObserver, 2000),
		);
	} else {
		setTimeout(setupObserver, 2000);
	}
}

/**
 * Observe the body to find when NotificationPane appears
 * @param {Electron.IpcRenderer} ipcRenderer
 */
function observeForNotificationPane(ipcRenderer) {
	const bodyObserver = new MutationObserver(() => {
		const notificationPane = document.querySelector(
			'[data-app-section="NotificationPane"]',
		);
		if (notificationPane) {
			console.log(
				"[Notification] NotificationPane appeared, setting up observer",
			);
			bodyObserver.disconnect();
			observeNotificationPane(notificationPane, ipcRenderer);
		}
	});

	if (document.body) {
		bodyObserver.observe(document.body, { childList: true, subtree: true });
	} else {
		setTimeout(() => observeForNotificationPane(ipcRenderer), 1000);
	}
}

/**
 * Observe the NotificationPane for new notifications
 * @param {Element} notificationPane
 * @param {Electron.IpcRenderer} ipcRenderer
 */
function observeNotificationPane(notificationPane, ipcRenderer) {
	const observer = new MutationObserver((mutations) => {
		mutations.forEach((mutation) => {
			mutation.addedNodes.forEach((node) => {
				if (node.nodeType === Node.ELEMENT_NODE) {
					processNotificationElement(node, ipcRenderer);
				}
			});
		});
	});

	observer.observe(notificationPane, { childList: true, subtree: true });
	console.log("[Notification] Observer attached to NotificationPane");
}

/**
 * Process a potential notification element
 * @param {Element} element
 * @param {Electron.IpcRenderer} ipcRenderer
 */
function processNotificationElement(element, ipcRenderer) {
	// 1. Check for reminder notifications (divs with timeuntildisplaystring attribute)
	const reminderDivs = element.querySelectorAll
		? [...element.querySelectorAll("[timeuntildisplaystring]")]
		: [];

	// Also check if the element itself has the attribute
	if (element.hasAttribute && element.hasAttribute("timeuntildisplaystring")) {
		reminderDivs.push(element);
	}

	reminderDivs.forEach((reminderDiv) => {
		const reminderData = extractReminderData(reminderDiv);
		if (reminderData) {
			console.log(
				"[Notification] Reminder notification detected:",
				reminderData,
			);
			ipcRenderer.invoke("showReminderNotification", reminderData);
		}
	});

	// 2. Check for email notifications (buttons with aria-label)
	const emailButtons = element.querySelectorAll
		? [
				...element.querySelectorAll("button[aria-label]"),
				...(element.matches?.("button[aria-label]") ? [element] : []),
			]
		: [];

	// Also check the element itself if it's a button
	if (element.tagName === "BUTTON" && element.hasAttribute("aria-label")) {
		emailButtons.push(element);
	}

	emailButtons.forEach((button) => {
		// Check if this is an email notification by DOM structure
		if (isEmailNotification(button)) {
			const ariaLabel = button.getAttribute("aria-label") || "";
			const emailData = extractEmailData(button, ariaLabel);
			if (emailData) {
				console.log("[Notification] Email notification detected:", emailData);
				ipcRenderer.invoke("showEmailNotification", emailData);
			}
		}
	});
}

/**
 * Check if button is an email notification by DOM structure
 * Email notifications have: .ZJg8d (sender), .KTZ84 (subject), .mrxI1 (body)
 * @param {Element} button
 * @returns {boolean}
 */
function isEmailNotification(button) {
	// Check for email notification structure
	const hasSender = button.querySelector(".ZJg8d");
	const hasSubject = button.querySelector(".KTZ84");
	const hasBody = button.querySelector(".mrxI1");

	return !!(hasSender && hasSubject && hasBody);
}

/**
 * Extract email data from notification button
 * @param {Element} button
 * @param {string} ariaLabel
 * @returns {{address: string, subject: string} | null}
 */
function extractEmailData(button, ariaLabel) {
	// 1. Get Sender (Outlook shows email if no display name)
	const senderElement = button.querySelector(".ZJg8d > div:first-child");
	const sender = senderElement?.textContent?.trim();

	// 2. Get Subject
	const subjectElement = button.querySelector(".KTZ84");
	const subject = subjectElement?.textContent?.trim();

	// 3. Extract Body
	const bodyElement = button.querySelector(".mrxI1");
	let messageBody = "";

	if (bodyElement) {
		const fullText = bodyElement.textContent;
		const emailRegex = /([a-zA-Z0-9._-]+@[a-zA-Z0-9._-]+\.[a-zA-Z0-9_-]+)/;

		// Process lines to clean up the body
		const lines = fullText.split("\n");
		const cleanLines = [];

		for (let i = 0; i < lines.length; i++) {
			const line = lines[i].trim();

			// Stop at email addresses (indicates reply footer starting)
			if (emailRegex.test(line)) break;

			// Stop at date patterns (e.g., "2026. jan. 24." or "Jan 24, 2026")
			if (/\d{4}\./.test(line) || /\d{1,2},\s*\d{4}/.test(line)) break;

			// Stop at common reply markers (lines ending with colon after name/time)
			if (/:\s*$/.test(line) && line.length < 50) break;

			// Add line if it's not empty and we haven't reached 3 lines yet
			if (line && cleanLines.length < 3) {
				cleanLines.push(line);
			} else if (cleanLines.length >= 3) {
				break;
			}
		}

		messageBody = cleanLines.join("\n");
	}

	// 4. Use sender from element or fallback to aria-label
	const formattedAddress =
		sender ||
		(() => {
			const colonIndex = ariaLabel.indexOf(":");
			return colonIndex > -1
				? ariaLabel.substring(colonIndex + 1).trim()
				: "Unknown";
		})();

	return {
		address: formattedAddress,
		subject: subject || "New message",
		body: messageBody,
	};
}

/**
 * Extract reminder data from reminder div element using DOM attributes
 * @param {Element} element - Element with timeuntildisplaystring attribute
 * @returns {{subject: string, location: string, timeUntil: string, startTime: string, reminderType: string} | null}
 */
function extractReminderData(element) {
	// Extract attributes directly from the element
	const subject = element.getAttribute("subject") || "";
	const location = element.getAttribute("location") || "";
	const timeUntil = element.getAttribute("timeuntildisplaystring") || "";
	const startTime = element.getAttribute("starttimedisplaystring") || "";
	const reminderType = element.getAttribute("remindertype") || "Reminder";

	// Return data if subject exists
	if (subject) {
		return {
			subject: subject,
			location: location,
			timeUntil: timeUntil,
			startTime: startTime,
			reminderType: reminderType,
		};
	}

	return null;
}

/**
 * Setup observer to watch Outlook's unread email counter
 * @param {Electron.IpcRenderer} ipcRenderer
 * @param {string} accountId
 */
function setupUnreadCountObserver(ipcRenderer, accountId) {
	try {
		let debounceTimer;
		const debouncedCheck = () => {
			clearTimeout(debounceTimer);
			debounceTimer = setTimeout(checkUnreadCount, 100);
		};

		const checkUnreadCount = () => {
			try {
				const unreadElements = document.querySelectorAll(".WIYG1.Mt2TB");
				let totalCount = 0;
				unreadElements.forEach((element) => {
					if (element.closest('[aria-labelledby="favoritesRoot"]')) {
						return;
					}
					const count = parseInt(element.textContent) || 0;
					totalCount += count;
				});

				console.log("[Unread Counter] Total unread count:", totalCount);

				if (ipcRenderer && ipcRenderer.invoke) {
					ipcRenderer.invoke("updateUnreadCount", {
						accountId,
						count: totalCount,
					});
				}
			} catch (e) {
				console.error("[Unread Counter] ERROR in checkUnreadCount:", e);
			}
		};

		const observer = new MutationObserver((mutations) => {
			try {
				let relevantMutation = false;
				for (const mutation of mutations) {
					// Check added nodes
					if (mutation.addedNodes) {
						mutation.addedNodes.forEach((node) => {
							// Ensure node is an element before accessing classList/querySelector
							if (node.nodeType === 1) {
								if (
									node.classList?.contains("WIYG1") ||
									(node.querySelector && node.querySelector(".WIYG1.Mt2TB"))
								) {
									relevantMutation = true;
								}
							}
						});
					}

					// Check removed nodes
					if (mutation.removedNodes) {
						mutation.removedNodes.forEach((node) => {
							if (node.nodeType === 1) {
								if (
									node.classList?.contains("WIYG1") ||
									(node.querySelector && node.querySelector(".WIYG1.Mt2TB"))
								) {
									relevantMutation = true;
								}
							}
						});
					}

					if (
						mutation.type === "characterData" &&
						mutation.target.parentElement?.classList.contains("WIYG1")
					) {
						relevantMutation = true;
					}
				}

				if (relevantMutation) {
					debouncedCheck();
				}
			} catch (e) {
				console.error("[Unread Counter] ERROR inside Observer callback:", e);
			}
		});

		// Defensive check: Does html element exist?
		const targetNode = document.documentElement || document;
		observer.observe(targetNode, {
			childList: true,
			subtree: true,
			characterData: true,
		});

		let attempts = 0;
		const maxAttempts = 20;

		const pollForElements = () => {
			console.log(`[Unread Counter] Polling attempt ${attempts + 1}...`);

			const unreadElements = document.querySelectorAll(".WIYG1.Mt2TB");

			if (unreadElements.length > 0) {
				console.log("[Unread Counter] Elements found!");
				checkUnreadCount();
			} else if (attempts < maxAttempts) {
				attempts++;
				setTimeout(pollForElements, 500);
			} else {
				console.log("[Unread Counter] Polling timeout reached.");
			}
		};

		setTimeout(pollForElements, 1000);
	} catch (e) {
		console.error("[Unread Counter] CRITICAL ERROR during setup:", e);
	}
}

/**
 * Setup observer to watch for active reminders
 * @param {Electron.IpcRenderer} ipcRenderer
 * @param {string} accountId
 */
function setupReminderCountObserver(ipcRenderer, accountId) {
	try {
		let debounceTimer;
		const debouncedCheck = () => {
			clearTimeout(debounceTimer);
			debounceTimer = setTimeout(checkReminderCount, 100);
		};

		const checkReminderCount = () => {
			try {
				// Count elements with timeuntildisplaystring attribute (active reminders)
				const reminderElements = document.querySelectorAll(
					"[timeuntildisplaystring]",
				);
				const totalCount = reminderElements.length;

				console.log("[Reminder Counter] Total reminder count:", totalCount);

				if (ipcRenderer && ipcRenderer.invoke) {
					console.log(
						"[Reminder Counter] Invoking updateReminderCount with count:",
						totalCount,
					);
					ipcRenderer
						.invoke("updateReminderCount", { accountId, count: totalCount })
						.then(() => {
							console.log(
								"[Reminder Counter] updateReminderCount invoke succeeded",
							);
						})
						.catch((err) => {
							console.error(
								"[Reminder Counter] updateReminderCount invoke FAILED:",
								err,
							);
						});
				} else {
					console.error(
						"[Reminder Counter] ipcRenderer or invoke not available!",
					);
				}
			} catch (e) {
				console.error("[Reminder Counter] ERROR in checkReminderCount:", e);
			}
		};

		const observer = new MutationObserver((mutations) => {
			try {
				let relevantMutation = false;
				for (const mutation of mutations) {
					// Check added/removed nodes for reminder elements
					if (mutation.addedNodes || mutation.removedNodes) {
						const nodeLists = [
							mutation.addedNodes,
							mutation.removedNodes,
						].filter(Boolean);
						nodeLists.forEach((nodeList) => {
							nodeList.forEach((node) => {
								if (node.nodeType === 1) {
									if (
										(node.hasAttribute &&
											node.hasAttribute("timeuntildisplaystring")) ||
										(node.querySelector &&
											node.querySelector("[timeuntildisplaystring]"))
									) {
										relevantMutation = true;
									}
								}
							});
						});
					}
				}

				if (relevantMutation) {
					debouncedCheck();
				}
			} catch (e) {
				console.error("[Reminder Counter] ERROR inside Observer callback:", e);
			}
		});

		// Observe the document for reminder changes
		const targetNode = document.documentElement || document;
		observer.observe(targetNode, {
			childList: true,
			subtree: true,
		});

		// Initial check after a delay
		setTimeout(checkReminderCount, 1000);
	} catch (e) {
		console.error("[Reminder Counter] CRITICAL ERROR during setup:", e);
	}
}

/**
 * Detect account email from page title
 * Outlook shows email in title like "Outlook - user@example.com"
 * @param {string} accountId
 * @param {Electron.IpcRenderer} ipcRenderer
 */
function detectAccountEmail(accountId, ipcRenderer) {
	let foundEmail = null;
	let emailSent = false;

	function checkEmail() {
		try {
			// First, try to find the primaryMailboxRoot element
			const primaryMailboxRoot = document.querySelector(
				'[id^="primaryMailboxRoot_"]',
			);
			if (primaryMailboxRoot) {
				// Get direct child spans only (not nested inside buttons or other elements)
				const emailSpan = Array.from(primaryMailboxRoot.children).find(
					(child) => child.tagName === "SPAN",
				);

				if (emailSpan) {
					// Get textContent from the direct child span
					const text = emailSpan.textContent?.trim() || "";

					// Check if it's an email address
					const emailMatch = text.match(/[\w.-]+@[\w.-]+\.\w+/);
					if (emailMatch) {
						const email = emailMatch[0];
						if (foundEmail !== email) {
							foundEmail = email;
							if (!emailSent) {
								emailSent = true;
								ipcRenderer.invoke("account-email-detected", {
									accountId,
									email,
								});
							}
							return;
						}
					}
				}
			}

			// Fallback: check page title
			const title = document.title || "";
			const emailMatch = title.match(/[\w.-]+@[\w.-]+\.\w+/);
			if (emailMatch) {
				const email = emailMatch[0];
				if (!emailSent) {
					emailSent = true;
					ipcRenderer.invoke("account-email-detected", { accountId, email });
				}
			}
		} catch (e) {
			console.error("[Account Email Detection] Error:", e);
		}
	}

	// Check once immediately
	checkEmail();

	// Watch for primaryMailboxRoot to appear (if page is still loading)
	// Wait for document.body to be available
	const startObserver = () => {
		if (!document.body) {
			// Try again in 100ms
			setTimeout(startObserver, 100);
			return;
		}

		const observer = new MutationObserver(() => {
			if (!emailSent) {
				checkEmail();
			}
		});

		observer.observe(document.body, { childList: true, subtree: true });

		// Stop observing after 30 seconds to save resources
		setTimeout(() => {
			observer.disconnect();
		}, 30000);
	};

	startObserver();
}

function initializeModules(config, ipcRenderer) {
	const zoomLevels = { '+': 0.25, '-': -0.25, '0': 0 };

	function saveZoomLevel() {
		ipcRenderer.invoke('saveZoomLevel', {
			partition: config.partition,
			zoomLevel: webFrame.getZoomLevel(),
		});
	}

	function setNextZoomLevel(keyName) {
		const zoomFactor = zoomLevels[keyName];
		if (typeof zoomFactor !== 'number') return;
		const currentZoom = webFrame.getZoomLevel();
		webFrame.setZoomLevel(zoomFactor === 0 ? 0 : currentZoom + zoomFactor);
		saveZoomLevel();
	}

	ipcRenderer.invoke('getZoomLevel', config.partition).then((zoomLevel) => {
		webFrame.setZoomLevel(zoomLevel);
	});
	ipcRenderer.on('zoom-changed', (_event, zoomDirection) => {
		setNextZoomLevel(zoomDirection === 'in' ? '+' : '-');
	});

	const keyMap = {
		'CTRL_+': () => setNextZoomLevel('+'),
		'CTRL_-': () => setNextZoomLevel('-'),
		CTRL_0: () => setNextZoomLevel('0'),
		ALT_ArrowLeft: () => window.history.back(),
		ALT_ArrowRight: () => window.history.forward(),
	};
	const handleKeyDown = (event) => {
		if (event.key === 'Control' || event.key === 'Alt') return;
		const key = `${event.ctrlKey ? 'CTRL_' : ''}${event.altKey ? 'ALT_' : ''}${event.key}`;
		if (typeof keyMap[key] === 'function') keyMap[key]();
	};
	const attachShortcuts = () => {
		window.addEventListener('keydown', handleKeyDown, false);
		const iframe = document.getElementsByTagName('iframe')[0];
		if (iframe?.contentDocument) {
			iframe.contentDocument.addEventListener('keydown', handleKeyDown, false);
		} else {
			setTimeout(attachShortcuts, 4000);
		}
	};
	if (document.readyState === 'loading') {
		document.addEventListener('DOMContentLoaded', attachShortcuts, { once: true });
	} else {
		attachShortcuts();
	}

	ipcRenderer.on('get-outlook-settings', (event) => {
		event.sender.send('get-outlook-settings', {});
	});
	ipcRenderer.on('set-outlook-settings', (event) => {
		event.sender.send('set-outlook-settings', true);
	});
}
