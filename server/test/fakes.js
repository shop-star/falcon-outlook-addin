/* In-memory stand-ins for Fergus and Claude, so the tests exercise the real
 * request/response handling without network access or API keys.
 */

export function fakeFergus({ customers = [], sites = [], failures = {} } = {}) {
  const calls = [];
  let nextId = 1000;

  async function fetchImpl(url, init = {}) {
    const { pathname, searchParams } = new URL(url);
    const method = init.method || "GET";
    const body = init.body ? JSON.parse(init.body) : undefined;
    calls.push({ method, path: pathname, query: Object.fromEntries(searchParams), body, auth: init.headers?.Authorization });

    const key = `${method} ${pathname}`;
    const failure = failures[key]?.shift?.();
    if (failure) {
      return new Response(JSON.stringify(failure.body ?? {}), {
        status: failure.status,
        headers: { "content-type": "application/json", ...(failure.headers || {}) },
      });
    }

    const json = (data, status = 200) =>
      new Response(JSON.stringify(data), { status, headers: { "content-type": "application/json" } });
    const text = (searchParams.get("filterSearchText") || "").toLowerCase();

    switch (key) {
      case "GET /customers":
        return json({ data: customers.filter((c) => JSON.stringify(c).toLowerCase().includes(text)), paging: {} });
      case "GET /sites":
        return json({ data: sites.filter((s) => JSON.stringify(s).toLowerCase().includes(text)), paging: {} });
      case "POST /customers":
        return json({ data: { id: nextId++, ...body } }, 201);
      case "POST /sites":
        return json({ data: { id: nextId++, ...body } }, 201);
      case "POST /jobs": {
        const id = nextId++;
        return json({ data: { id, jobNo: `J${id}`, ...body } }, 201);
      }
      case "POST /notes":
      case "POST /calendarEvents":
        return json({ data: { id: nextId++, ...body } }, 201);
      default:
        return json({ message: "not found" }, 404);
    }
  }

  return { fetchImpl, calls };
}

export function fakeAnthropic(extraction) {
  const requests = [];
  return {
    requests,
    beta: {
      messages: {
        async create(params) {
          requests.push(params);
          return {
            stop_reason: "end_turn",
            content: [{ type: "text", text: JSON.stringify(extraction) }],
          };
        },
      },
    },
  };
}

const f = (value, confidence = "high") => ({ value, confidence });

export const sampleExtraction = {
  senderRole: "builder_or_agent",
  customerName: f("Aroha Ngata"),
  company: f(""),
  contactName: f("Aroha Ngata"),
  phone: f("021 555 0142"),
  email: f("aroha.ngata@example.com"),
  siteAddress: { address1: "14 Kowhai Road", address2: "", suburb: "Mt Eden", city: "Auckland", region: "", postcode: "1024", confidence: "high" },
  jobTitle: f("Decking timber – 14 Kowhai Rd"),
  jobDescription: f("Supply 40 m² of 140x32 kwila decking with fixings."),
  group: f("Decking"),
  notes: "Sent by the builder on behalf of the homeowner.",
};

export const sampleEmail = {
  internetMessageId: "<CAF123@mail.example.com>",
  subject: "Quote please – deck at 14 Kowhai Rd",
  from: { name: "Sam Builder", address: "sam@builders.example.com" },
  receivedAt: "2026-10-12T21:15:00.000Z",
  bodyText: "Hi team, my client Aroha Ngata (021 555 0142, aroha.ngata@example.com) needs 40m2 of 140x32 kwila decking delivered to 14 Kowhai Road, Mt Eden, Auckland 1024. Cheers, Sam",
  attachments: [{ name: "deck-plan.pdf", size: 120000, contentType: "application/pdf" }],
};
