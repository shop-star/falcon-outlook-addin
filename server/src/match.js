/* Find existing Fergus customers and sites that might be the ones this email
 * is about.
 *
 * Fergus only offers a free-text search (filterSearchText) on customers and
 * sites, so each extracted detail (email, phone, name, company) is searched
 * separately and the results are ranked by how many of them a record matched
 * — a customer found by both email and phone outranks one found by name alone.
 */

const MAX_CANDIDATES = 5;

function norm(s) {
  return (s || "").toLowerCase().replace(/\s+/g, " ").trim();
}

function digits(s) {
  return (s || "").replace(/\D/g, "");
}

// NZ numbers are written as both 021 123 4567 and +64 21 123 4567.
function samePhone(a, b) {
  const x = digits(a).replace(/^64/, "0");
  const y = digits(b).replace(/^64/, "0");
  return x.length >= 7 && x === y;
}

async function settle(promise) {
  try {
    return { ok: true, value: await promise };
  } catch (error) {
    return { ok: false, error };
  }
}

export async function findCandidates(fergus, extraction) {
  const v = (field) => extraction?.[field]?.value || "";
  const email = v("email");
  const phone = v("phone");
  const names = [...new Set([v("customerName"), v("company"), v("contactName")].filter(Boolean))];
  const site = extraction?.siteAddress || {};

  const customerQueries = [];
  if (email) customerQueries.push({ by: "email", text: email });
  if (phone) customerQueries.push({ by: "phone", text: phone });
  for (const name of names) customerQueries.push({ by: "name", text: name });

  const siteQueries = [];
  if (site.address1) siteQueries.push({ by: "address", text: site.address1, postcode: site.postcode || "" });

  const [customerResults, siteResults] = await Promise.all([
    Promise.all(customerQueries.map((q) => settle(fergus.searchCustomers(q.text)).then((r) => ({ ...r, q })))),
    Promise.all(siteQueries.map((q) => settle(fergus.searchSites(q.text, q.postcode)).then((r) => ({ ...r, q })))),
  ]);

  // A failed search shouldn't stop the person creating a job — they can
  // still pick "new" — but a rate limit or bad token is worth telling them.
  const errors = [...customerResults, ...siteResults].filter((r) => !r.ok).map((r) => r.error);

  const customers = rank(customerResults, (c, q) => {
    if (q.by === "email") return norm(c.email) === norm(q.text) ? 3 : 1;
    if (q.by === "phone") return samePhone(c.phone, q.text) ? 3 : 1;
    return norm(c.name) === norm(q.text) || norm(c.contactName) === norm(q.text) ? 2 : 1;
  });

  const sites = rank(siteResults, (s, q) => {
    const sameStreet = norm(s.address).startsWith(norm(q.text));
    const samePostcode = q.postcode && s.postcode && q.postcode === s.postcode;
    return (sameStreet ? 2 : 1) + (samePostcode ? 1 : 0);
  });

  return { customers, sites, errors };
}

function rank(results, score) {
  const byId = new Map();
  for (const r of results) {
    if (!r.ok) continue;
    for (const record of r.value) {
      if (record.id == null) continue;
      const entry = byId.get(record.id) || { ...record, score: 0, matchedOn: [] };
      entry.score += score(record, r.q);
      if (!entry.matchedOn.includes(r.q.by)) entry.matchedOn.push(r.q.by);
      byId.set(record.id, entry);
    }
  }
  return [...byId.values()].sort((a, b) => b.score - a.score).slice(0, MAX_CANDIDATES);
}
