/* global localStorage, fetch, console */

/**
 * LicenseGate – license validation module
 *
 * Validates license keys against the LicenseGate REST API.
 * Caches a validated key in localStorage so the user doesn't have to
 * re-enter it every session.
 */

const LICENSEGATE_USER_ID = "a241b";
const LICENSEGATE_API = "https://api.licensegate.io/license";
const LS_KEY = "licensegate_key";

/**
 * Validate a license key against the LicenseGate API.
 * @param {string} licenseKey
 * @returns {Promise<{valid: boolean, status: string}>}
 */
async function validateKey(licenseKey) {
    try {
        const url = `${LICENSEGATE_API}/${LICENSEGATE_USER_ID}/${encodeURIComponent(licenseKey)}/verify`;
        const response = await fetch(url, {
            method: "GET",
            headers: { Accept: "application/json" },
        });

        if (!response.ok) {
            return { valid: false, status: "SERVER_ERROR" };
        }

        const data = await response.json();

        // LicenseGate returns { result: "VALID" } on success
        const isValid =
            data.valid === true ||
            (data.result && data.result.toUpperCase() === "VALID");

        return { valid: isValid, status: data.result || (isValid ? "VALID" : "INVALID") };
    } catch (err) {
        console.error("LicenseGate validation error:", err);
        return { valid: false, status: "NETWORK_ERROR" };
    }
}

/**
 * Check whether a previously cached license key is still valid.
 * @returns {Promise<boolean>}
 */
export async function checkLicense() {
    try {
        const key = localStorage.getItem(LS_KEY);
        if (!key) return false;

        const { valid } = await validateKey(key);
        if (!valid) {
            localStorage.removeItem(LS_KEY);
        }
        return valid;
    } catch {
        return false;
    }
}

/**
 * Activate (validate + cache) a new license key.
 * @param {string} licenseKey
 * @returns {Promise<{valid: boolean, status: string}>}
 */
export async function activateLicense(licenseKey) {
    const trimmed = (licenseKey || "").trim();
    if (!trimmed) return { valid: false, status: "EMPTY_KEY" };

    const result = await validateKey(trimmed);
    if (result.valid) {
        localStorage.setItem(LS_KEY, trimmed);
    }
    return result;
}

/**
 * Remove the cached license key (deactivate).
 */
export function clearLicense() {
    localStorage.removeItem(LS_KEY);
}

/**
 * Return the currently cached license key (or empty string).
 */
export function getCachedKey() {
    try {
        return localStorage.getItem(LS_KEY) || "";
    } catch {
        return "";
    }
}
