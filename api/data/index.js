const { TableClient, AzureNamedKeyCredential } = require("@azure/data-tables");

const CONNECTION_STRING = process.env.AZURE_STORAGE_CONNECTION_STRING;

async function getClient(table) {
  const client = TableClient.fromConnectionString(CONNECTION_STRING, table);
  try { await client.createTable(); } catch(e) {}
  return client;
}

module.exports = async function(context, req) {
  const table = req.query.table;
  const method = req.method;

  if (!table) {
    context.res = { status: 400, body: "table required" };
    return;
  }

  try {
    const client = await getClient(table);

    if (method === "GET") {
      const items = [];
      for await (const entity of client.listEntities()) {
        items.push(entity);
      }
      context.res = { body: items };

    } else if (method === "POST") {
      const item = req.body;
      await client.upsertEntity({
        partitionKey: "main",
        rowKey: item.id,
        data: JSON.stringify(item)
      }, "Replace");
      context.res = { body: { ok: true } };

    } else if (method === "DELETE") {
      const id = req.query.id;
      await client.deleteEntity("main", id);
      context.res = { body: { ok: true } };
    }

  } catch(e) {
    context.res = { status: 500, body: e.message };
  }
};
