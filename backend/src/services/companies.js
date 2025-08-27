const odbc = require('odbc');

// Connection string provided by user (added CHARSET for proper encoding)
const CONN_STR = (
  "DRIVER={SQL Anywhere 17};"
  + "HOST=NOTE-GO-273.go.local:2638;"
  + "DBN=contabil;"
  + "UID=ESTATISTICA002;"
  + "PWD=U0T/wq6OdZ0oYSpvJRWGfg==;"
  + "CHARSET=UTF8;"
);

async function getCompanies(searchTerm){
  let connection;
  try{
    connection = await odbc.connect(CONN_STR);
    let result;
    if(searchTerm && String(searchTerm).trim().length>0){
      const like = `%${searchTerm}%`;
      const sql = `select g.apel_emp as apel_emp, g.codi_emp as codi_emp from bethadba.geempre g where UPPER(g.apel_emp) LIKE UPPER(?) OR CAST(g.codi_emp AS VARCHAR(50)) LIKE ?`;
      result = await connection.query(sql, [like, like]);
    } else {
      const sql = `select g.apel_emp as apel_emp, g.codi_emp as codi_emp from bethadba.geempre g`;
      result = await connection.query(sql);
    }
    // Normalize column names to { code, name }
    return (result || []).map(r=>({
      code: r.CODI_EMP ?? r.codi_emp ?? r.code ?? null,
      name: r.APEL_EMP ?? r.apel_emp ?? r.name ?? null
    }));
  }catch(err){
    throw err;
  }finally{
    if(connection){ try{ await connection.close(); }catch(e){} }
  }
}

module.exports = { getCompanies };
