// Is the running deployment's baked copy of the user list out of date?
//
// Vercel bakes env vars into a deployment at BUILD time. loadUsers() reads
// DFE_USERS_JSON out of process.env, so the running functions keep serving the
// value as it was when they were built. Saving a user PATCHes the env var, but
// the running deployment cannot see it.
//
// That is what made creating two users in a row lose the first: the second
// create read the stale baked list (without user one), pushed it back with user
// two appended, and user one was gone. The same shape applies to an edit or a
// delete done before the redeploy lands.
//
// So: compare the env var's updatedAt against the timestamp baked in at build.
// If the env var moved after this build, the baked list is stale and no further
// write may be made until a new deployment is serving. Only metadata is read,
// which keeps working even if DFE_USERS_JSON is ever marked sensitive (a
// sensitive var reads back blank through the API).

const PROJECT_ID = 'prj_FaBoeZxXminOA9W8gSwsrwuLTz2i';
const TEAM_ID = 'team_CUmgfjtVYHnIgiFo3lIvucqh';
const VERCEL_KEY = 'DFE_USERS_JSON';

export interface DeployState {
  // true = a user write is saved but not yet live; further writes are blocked.
  pending: boolean;
  // true = a production build is running right now, so pending will clear itself.
  building: boolean;
  // Set when we could not determine the state. Writes are blocked on unknown:
  // guessing wrong here silently deletes a user record.
  unknown: boolean;
  savedAt: string | null;   // when the user list was last written
  builtAt: string | null;   // when the running deployment baked its copy
  message: string;
}

function api(path: string): string {
  const sep = path.includes('?') ? '&' : '?';
  return `https://api.vercel.com${path}${sep}teamId=${TEAM_ID}`;
}

async function vercelGet(path: string, token: string) {
  const res = await fetch(api(path), {
    headers: { Authorization: `Bearer ${token}` },
    cache: 'no-store',
  });
  if (!res.ok) throw new Error(`${path} -> ${res.status}`);
  return res.json();
}

export async function getDeployState(): Promise<DeployState> {
  const base: DeployState = {
    pending: false, building: false, unknown: false,
    savedAt: null, builtAt: null, message: '',
  };

  // Off Vercel (local dev) the file is the store and reads are always live.
  if (!process.env.VERCEL) return { ...base, message: 'Local dev — changes are live immediately.' };

  const token = process.env.VERCEL_TOKEN;
  const builtAt = process.env.DFE_BUILD_TIME || null;

  if (!token) {
    return { ...base, unknown: true, builtAt,
      message: 'VERCEL_TOKEN is not set, so the app cannot tell whether saved users are live yet.' };
  }
  if (!builtAt) {
    return { ...base, unknown: true,
      message: 'This deployment has no build timestamp, so the app cannot tell whether saved users are live yet.' };
  }

  try {
    const { envs } = await vercelGet(`/v9/projects/${PROJECT_ID}/env`, token) as {
      envs: { key: string; updatedAt?: number; createdAt?: number }[];
    };
    const record = envs.find(e => e.key === VERCEL_KEY);
    if (!record) {
      return { ...base, unknown: true, builtAt,
        message: `${VERCEL_KEY} was not found on this project.` };
    }

    const savedMs = record.updatedAt ?? record.createdAt ?? 0;
    const builtMs = Date.parse(builtAt);
    const pending = savedMs > builtMs;
    const savedAt = savedMs ? new Date(savedMs).toISOString() : null;

    if (!pending) {
      return { ...base, savedAt, builtAt, message: 'All saved users are live.' };
    }

    // Pending — is a build already on its way?
    let building = false;
    try {
      const { deployments } = await vercelGet(
        `/v6/deployments?projectId=${PROJECT_ID}&target=production&limit=5`, token,
      ) as { deployments: { state?: string; readyState?: string; created?: number }[] };
      building = (deployments || []).some(d => {
        const state = d.readyState || d.state;
        return state === 'BUILDING' || state === 'QUEUED' || state === 'INITIALIZING';
      });
    } catch (err) {
      console.error('[deployState] deployment list failed:', err);
    }

    return {
      ...base, pending: true, building, savedAt, builtAt,
      message: building
        ? 'A user was saved and is going live now. This clears itself when the deploy finishes.'
        : 'A user was saved but is not live yet. It needs a deploy before anything else can be changed.',
    };
  } catch (err) {
    console.error('[deployState] check failed:', err);
    return { ...base, unknown: true, builtAt,
      message: 'Could not reach Vercel to check whether saved users are live yet.' };
  }
}

// The deployment to rebuild. VERCEL_DEPLOYMENT_ID is a system env var and is
// only injected when the project exposes those, so fall back to looking up the
// newest ready production deployment. Without the fallback both the automatic
// trigger and the manual "Deploy now" button would fail on a project with
// system variables turned off, which would leave an admin locked out of user
// management with no way forward.
async function currentDeploymentId(token: string): Promise<string | null> {
  if (process.env.VERCEL_DEPLOYMENT_ID) return process.env.VERCEL_DEPLOYMENT_ID;
  try {
    const { deployments } = await vercelGet(
      `/v6/deployments?projectId=${PROJECT_ID}&target=production&state=READY&limit=1`, token,
    ) as { deployments: { uid?: string; id?: string }[] };
    const newest = (deployments || [])[0];
    return newest?.uid || newest?.id || null;
  } catch (err) {
    console.error('[deployState] could not find a deployment to rebuild:', err);
    return null;
  }
}

// Redeploy the currently-running production deployment. Same commit, rebuilt,
// so it bakes in the current DFE_USERS_JSON. Returns an error string on failure.
export async function triggerRedeploy(): Promise<{ ok: boolean; error?: string }> {
  const token = process.env.VERCEL_TOKEN;
  if (!token) return { ok: false, error: 'VERCEL_TOKEN is not set' };

  const deploymentId = await currentDeploymentId(token);
  if (!deploymentId) return { ok: false, error: 'Could not find a deployment to rebuild' };

  try {
    const res = await fetch(api('/v13/deployments?forceNew=1&skipAutoDetectionConfirmation=1'), {
      method: 'POST',
      headers: { Authorization: `Bearer ${token}`, 'Content-Type': 'application/json' },
      body: JSON.stringify({ name: 'defy-field-execute', deploymentId, target: 'production' }),
    });
    if (!res.ok) {
      const detail = await res.text();
      console.error('[deployState] redeploy failed:', res.status, detail.slice(0, 500));
      return { ok: false, error: `Vercel returned ${res.status}` };
    }
    return { ok: true };
  } catch (err) {
    console.error('[deployState] redeploy threw:', err);
    return { ok: false, error: String(err) };
  }
}
