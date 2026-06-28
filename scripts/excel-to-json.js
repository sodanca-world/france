import fs from 'node:fs';
import path from 'node:path';
import xlsx from 'xlsx';

const excelPath = path.resolve(process.cwd(), 'sodanca-france-airtable-import.xlsx');
const outputDir = path.resolve(process.cwd(), 'src/content');

if (!fs.existsSync(outputDir)) {
  fs.mkdirSync(outputDir, { recursive: true });
}

console.log('Lendo planilha Excel:', excelPath);
if (!fs.existsSync(excelPath)) {
  console.error('Erro: Arquivo Excel nao encontrado em:', excelPath);
  process.exit(1);
}

try {
  const workbook = xlsx.readFile(excelPath);

  // 1. Processar Cidades
  const sheetCidades = workbook.Sheets['Cidades'];
  if (!sheetCidades) {
    throw new Error('Aba "Cidades" nao encontrada na planilha.');
  }
  const cidadesData = xlsx.utils.sheet_to_json(sheetCidades);
  const cidadesMapped = cidadesData.map(row => ({
    id: row.Slug || String(row.Nome).toLowerCase().replace(/\s+/g, '-'),
    Nome: row.Nome,
    Slug: row.Slug,
    Ordem: row.Ordem ? Number(row.Ordem) : undefined
  }));

  fs.writeFileSync(
    path.join(outputDir, 'cidades.json'),
    JSON.stringify(cidadesMapped, null, 2),
    'utf-8'
  );
  console.log(`Salvas ${cidadesMapped.length} cidades em src/content/cidades.json`);

  // 2. Processar Lojas
  const sheetLojas = workbook.Sheets['Lojas'];
  if (!sheetLojas) {
    throw new Error('Aba "Lojas" nao encontrada na planilha.');
  }
  const lojasData = xlsx.utils.sheet_to_json(sheetLojas);
  const lojasMapped = lojasData.map(row => {
    const cidade = row['Cidade (slug)'];
    // Cidade deve ser um array de strings contendo o slug/id da cidade correspondente
    const cidadesArray = cidade ? [cidade] : [];
    
    // Tratamento do Logo
    let logoArray = undefined;
    const logoFile = row['Logo (URL)'];
    if (logoFile) {
      logoArray = [{ url: `/assets/images/${logoFile}` }];
    }

    return {
      id: row.Slug,
      Nome: row.Nome,
      Slug: row.Slug,
      Cidade: cidadesArray,
      Endereco: row.Endereco || undefined,
      LinkMaps: row.LinkMaps || undefined,
      Telefone: row.Telefone ? String(row.Telefone) : undefined,
      Email: row.Email || undefined,
      Instagram: row.Instagram || undefined,
      Facebook: row.Facebook || undefined,
      Website: row.Website || undefined,
      Logo: logoArray
    };
  });

  fs.writeFileSync(
    path.join(outputDir, 'lojas.json'),
    JSON.stringify(lojasMapped, null, 2),
    'utf-8'
  );
  console.log(`Salvas ${lojasMapped.length} lojas em src/content/lojas.json`);
} catch (error) {
  console.error('Erro ao ler a planilha Excel:', error);
  process.exit(1);
}
