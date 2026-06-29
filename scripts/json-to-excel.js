import fs from 'node:fs';
import path from 'node:path';
import xlsx from 'xlsx';

const inputDir = path.resolve(process.cwd(), 'src/content');
const excelPath = path.resolve(process.cwd(), 'sodanca-france-airtable-import.xlsx');

const cidadesJsonPath = path.join(inputDir, 'cidades.json');
const lojasJsonPath = path.join(inputDir, 'lojas.json');

console.log('Lendo dados JSON atuais...');

if (!fs.existsSync(cidadesJsonPath) || !fs.existsSync(lojasJsonPath)) {
  console.error('Erro: Arquivos json de cidades ou lojas nao encontrados.');
  process.exit(1);
}

try {
  const cidades = JSON.parse(fs.readFileSync(cidadesJsonPath, 'utf-8'));
  const lojas = JSON.parse(fs.readFileSync(lojasJsonPath, 'utf-8'));

  // 1. Preparar dados para aba Cidades
  const cidadesRows = cidades.map((c, index) => ({
    Nome: c.Nome,
    Slug: c.Slug,
    Ordem: c.Ordem || (index + 1)
  }));

  // 2. Preparar dados para aba Lojas
  const lojasRows = lojas.map(l => {
    // Pegar o slug da cidade (loja.Cidades é um array, ex: ["paris"])
    const cidadeSlug = l.Cidades && l.Cidades.length > 0 ? l.Cidades[0] : '';
    
    // Extrair o nome do arquivo do logo do array (ex: /assets/images/flashdance.jpg -> flashdance.jpg)
    let logoFile = '';
    if (l.Logo && l.Logo.length > 0) {
      const url = l.Logo[0].url;
      logoFile = url.replace('/assets/images/', '');
    }

    return {
      Nome: l.Nome,
      Cidade: l.Cidade || '',
      Slug: l.Slug,
      'Cidade (slug)': cidadeSlug,
      Endereco: l.Endereco || '',
      LinkMaps: l.LinkMaps || '',
      Telefone: l.Telefone || '',
      Email: l.Email || '',
      Instagram: l.Instagram || '',
      Facebook: l.Facebook || '',
      Website: l.Website || '',
      'Logo (URL)': logoFile,
      Ordem: l.Ordem || ''
    };
  });

  // Criar WorkBook do Excel
  const workbook = xlsx.utils.book_new();

  // Criar planilha Cidades
  const sheetCidades = xlsx.utils.json_to_sheet(cidadesRows);
  xlsx.utils.book_append_sheet(workbook, sheetCidades, 'Cidades');

  // Criar planilha Lojas
  const sheetLojas = xlsx.utils.json_to_sheet(lojasRows);
  xlsx.utils.book_append_sheet(workbook, sheetLojas, 'Lojas');

  // Gravar arquivo Excel
  xlsx.writeFile(workbook, excelPath);
  console.log(`=== Planilha gerada com sucesso em: ${excelPath} ===`);
  console.log(`Salvos ${cidadesRows.length} cidades e ${lojasRows.length} lojas.`);
} catch (error) {
  console.error('Erro ao gerar a planilha Excel:', error);
  process.exit(1);
}
