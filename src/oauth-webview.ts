import { spawnSync } from "child_process";
import { createHash, randomBytes } from "crypto";

const REDIRECT_URI =
  "https://login.microsoftonline.com/common/oauth2/nativeclient";
const LOGIN_TIMEOUT_MS = 10 * 60 * 1000;

const WEBVIEW_SCRIPT = String.raw`
ObjC.import("Cocoa");
ObjC.import("WebKit");

let mscliApp;
let mscliWindow;
let mscliWebView;
let mscliTimer;
let mscliTimerTarget;
let mscliCallbackPrefix;
let mscliStartedAt;
let mscliTimeoutMs;
let mscliFinished = false;

function finish(result) {
  if (mscliFinished) return;
  mscliFinished = true;
  console.log(result);
  if (mscliTimer) mscliTimer.invalidate;
  if (mscliWindow) mscliWindow.close;
  mscliApp.terminate(null);
}

ObjC.registerSubclass({
  name: "MsCliOAuthTimerTarget",
  methods: {
    "tick:": {
      types: ["void", ["id"]],
      implementation: function () {
        if (mscliWebView.URL) {
          const currentUrl = ObjC.unwrap(mscliWebView.URL.absoluteString);
          if (currentUrl.toLowerCase().startsWith(mscliCallbackPrefix)) {
            mscliWebView.stopLoading;
            finish("MSCLI_CALLBACK:" + currentUrl);
            return;
          }
        }

        const elapsed = Date.now() - mscliStartedAt;
        if (elapsed > 500 && !ObjC.unwrap(mscliWindow.visible)) {
          finish("MSCLI_CANCELLED");
          return;
        }
        if (elapsed >= mscliTimeoutMs) finish("MSCLI_TIMEOUT");
      },
    },
  },
});

function installStandardMenus() {
  const mainMenu = $.NSMenu.alloc.initWithTitle("MainMenu");

  const appMenuItem = $.NSMenuItem.alloc.init;
  const appMenu = $.NSMenu.alloc.initWithTitle("ms-cli");
  appMenu.addItemWithTitleActionKeyEquivalent(
    "Quit ms-cli",
    "terminate:",
    "q"
  );
  appMenuItem.submenu = appMenu;
  mainMenu.addItem(appMenuItem);

  const editMenuItem = $.NSMenuItem.alloc.init;
  const editMenu = $.NSMenu.alloc.initWithTitle("Edit");
  editMenu.addItemWithTitleActionKeyEquivalent("Cut", "cut:", "x");
  editMenu.addItemWithTitleActionKeyEquivalent("Copy", "copy:", "c");
  editMenu.addItemWithTitleActionKeyEquivalent("Paste", "paste:", "v");
  editMenu.addItem($.NSMenuItem.separatorItem);
  editMenu.addItemWithTitleActionKeyEquivalent(
    "Select All",
    "selectAll:",
    "a"
  );
  editMenuItem.submenu = editMenu;
  mainMenu.addItem(editMenuItem);

  mscliApp.mainMenu = mainMenu;
}

function run(argv) {
  const authorizationUrl = argv[0];
  mscliCallbackPrefix = argv[1].toLowerCase();
  mscliTimeoutMs = Number(argv[2]);

  mscliApp = $.NSApplication.sharedApplication;
  mscliApp.setActivationPolicy($.NSApplicationActivationPolicyRegular);
  mscliApp.finishLaunching;
  installStandardMenus();

  const style =
    $.NSWindowStyleMaskTitled |
    $.NSWindowStyleMaskClosable |
    $.NSWindowStyleMaskResizable;
  mscliWindow = $.NSWindow.alloc.initWithContentRectStyleMaskBackingDefer(
    $.NSMakeRect(0, 0, 540, 720),
    style,
    $.NSBackingStoreBuffered,
    false
  );
  mscliWindow.title = "ms-cli Microsoft 365 login";

  const configuration = $.WKWebViewConfiguration.alloc.init;
  configuration.websiteDataStore = $.WKWebsiteDataStore.defaultDataStore;
  mscliWebView = $.WKWebView.alloc.initWithFrameConfiguration(
    mscliWindow.contentView.bounds,
    configuration
  );
  mscliWebView.autoresizingMask =
    $.NSViewWidthSizable | $.NSViewHeightSizable;
  mscliWindow.contentView.addSubview(mscliWebView);

  mscliWindow.center;
  mscliWindow.makeKeyAndOrderFront(null);
  mscliWindow.orderFrontRegardless;
  mscliApp.activateIgnoringOtherApps(true);

  const request = $.NSURLRequest.requestWithURL(
    $.NSURL.URLWithString(authorizationUrl)
  );
  mscliWebView.loadRequest(request);

  mscliStartedAt = Date.now();
  mscliTimerTarget = $.MsCliOAuthTimerTarget.alloc.init;
  mscliTimer = $.NSTimer.scheduledTimerWithTimeIntervalTargetSelectorUserInfoRepeats(
    0.05,
    mscliTimerTarget,
    "tick:",
    null,
    true
  );
  mscliApp.run;
}
`;

interface OAuthTokenResponse {
  access_token?: string;
  refresh_token?: string;
  id_token?: string;
  expires_in?: number;
  error?: string;
  error_description?: string;
}

export interface InteractiveTokenResult {
  accessToken: string;
  refreshToken: string;
  idToken?: string;
  expiresIn?: number;
}

function randomBase64Url(bytes = 48): string {
  return randomBytes(bytes).toString("base64url");
}

function openLoginWebView(authorizationUrl: string): URL {
  const result = spawnSync(
    "/usr/bin/osascript",
    [
      "-l",
      "JavaScript",
      "-e",
      WEBVIEW_SCRIPT,
      authorizationUrl,
      REDIRECT_URI,
      String(LOGIN_TIMEOUT_MS),
    ],
    {
      encoding: "utf8",
      maxBuffer: 1024 * 1024,
      timeout: LOGIN_TIMEOUT_MS + 15_000,
      stdio: ["ignore", "pipe", "pipe"],
    }
  );

  const lines = `${result.stdout}\n${result.stderr}`
    .split(/\r?\n/)
    .map((line) => line.trim());
  if (lines.includes("MSCLI_CANCELLED")) {
    throw new Error("Microsoft 365 のログインがキャンセルされました。");
  }
  if (lines.includes("MSCLI_TIMEOUT")) {
    throw new Error("Microsoft 365 のログインがタイムアウトしました。");
  }
  const marker = "MSCLI_CALLBACK:";
  const callback = lines.find((line) => line.startsWith(marker));
  if (result.error || result.status !== 0 || !callback) {
    throw new Error("Microsoft 365 から認証結果を受け取れませんでした。");
  }
  return new URL(callback.slice(marker.length));
}

/**
 * Authenticate in a persistent macOS WebView using authorization-code + PKCE.
 * Device Code Flow and a client secret are not used.
 */
export async function acquireInteractiveToken(
  clientId: string,
  tenantId: string,
  scope: string
): Promise<InteractiveTokenResult> {
  const verifier = randomBase64Url();
  const challenge = createHash("sha256")
    .update(verifier)
    .digest("base64url");
  const state = randomBase64Url(24);

  const authorizeUrl = new URL(
    `https://login.microsoftonline.com/${encodeURIComponent(tenantId)}/oauth2/v2.0/authorize`
  );
  authorizeUrl.search = new URLSearchParams({
    client_id: clientId,
    response_type: "code",
    redirect_uri: REDIRECT_URI,
    response_mode: "query",
    scope,
    code_challenge: challenge,
    code_challenge_method: "S256",
    state,
    prompt: "select_account",
  }).toString();

  const callback = openLoginWebView(authorizeUrl.toString());
  const oauthError = callback.searchParams.get("error");
  if (oauthError) {
    const description = callback.searchParams.get("error_description");
    throw new Error(
      `Microsoft 365 login failed: ${description ?? oauthError}`
    );
  }
  if (callback.searchParams.get("state") !== state) {
    throw new Error("Microsoft 365 login failed: OAuth state mismatch");
  }
  const code = callback.searchParams.get("code");
  if (!code) {
    throw new Error("Microsoft 365 login failed: authorization code is missing");
  }

  const response = await fetch(
    `https://login.microsoftonline.com/${encodeURIComponent(tenantId)}/oauth2/v2.0/token`,
    {
      method: "POST",
      headers: { "Content-Type": "application/x-www-form-urlencoded" },
      body: new URLSearchParams({
        client_id: clientId,
        grant_type: "authorization_code",
        code,
        redirect_uri: REDIRECT_URI,
        code_verifier: verifier,
        scope,
      }),
    }
  );
  const token = (await response.json()) as OAuthTokenResponse;
  if (!response.ok || token.error) {
    throw new Error(
      `Microsoft 365 token exchange failed (${response.status}): ` +
        (token.error_description ?? token.error ?? "unknown error")
    );
  }
  if (!token.access_token || !token.refresh_token) {
    throw new Error(
      "Microsoft 365 login succeeded, but access/refresh token was not issued."
    );
  }

  return {
    accessToken: token.access_token,
    refreshToken: token.refresh_token,
    idToken: token.id_token,
    expiresIn: token.expires_in,
  };
}
