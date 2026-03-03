import { execSync } from "child_process";

/**
 * Require Touch ID authentication before proceeding.
 * Throws if authentication fails or Touch ID is unavailable.
 */
export function requireTouchId(reason: string): void {
  const script = `
ObjC.import("LocalAuthentication");
ObjC.import("Foundation");
var ctx = $.LAContext.alloc.init;
var err = Ref();
if (!ctx.canEvaluatePolicyError(2, err)) {
  throw "Authentication not available";
}
var done = false;
var authOk = false;
ctx.evaluatePolicyLocalizedReasonReply(2, ${JSON.stringify(reason)}, function(success, error) {
  authOk = success;
  done = true;
});
while (!done) {
  $.NSRunLoop.currentRunLoop.runUntilDate($.NSDate.dateWithTimeIntervalSinceNow(0.1));
}
if (!authOk) throw "Auth failed";
"OK";
`;

  try {
    execSync(`osascript -l JavaScript -e '${script.replace(/'/g, "'\"'\"'")}'`, {
      stdio: ["pipe", "pipe", "pipe"],
      timeout: 60000,
    });
  } catch {
    console.error("Touch ID authentication failed or cancelled.");
    process.exit(1);
  }
}
