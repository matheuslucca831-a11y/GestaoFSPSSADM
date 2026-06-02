'use client';
import ExcelJS from 'exceljs';

import React, { useState, useEffect, useRef } from 'react';
import { 
  Save, Building2, CheckCircle2, Phone, Download, Upload, RefreshCw 
} from 'lucide-react';
import { db } from '../../db';

export default function Configuracoes() {
  const [unidade, setUnidade] = useState('');
  const [telefone, setTelefone] = useState('');
  const [salvo, setSalvo] = useState(false);
  const [statusBackup, setStatusBackup] = useState<{ tipo: 'sucesso' | 'erro'; msg: string } | null>(null);
  const fileInputRef = useRef<HTMLInputElement>(null);

  const handleExportarExcel = async () => {
    try {
      setStatusBackup({ tipo: 'sucesso', msg: 'Gerando relatório completo...' });
      
      const workbook = new ExcelJS.Workbook();
      
      // Estilo de cabeçalho padrão com correção de Tipo para o ExcelJS
      const headerStyle = {
        font: { bold: true, color: { argb: 'FFFFFFFF' } },
        fill: { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FF2563EB' } } as const, // <-- ADICIONADO "as const" AQUI
        alignment: { horizontal: 'center' }
      };
  
      // Função auxiliar para criar abas
      const criarAba = async (nomeAba: string, tabela: string, colunas: any[]) => {
        const ws = workbook.addWorksheet(nomeAba);
        ws.columns = colunas;
        ws.getRow(1).eachCell((cell) => {
          cell.fill = headerStyle.fill; 
          cell.font = headerStyle.font; 
        });
        const dados = await db.table(tabela).toArray();
        dados.forEach(item => ws.addRow(item));
      };
  
      // 1. ABA: ENCAMINHAMENTOS (Legado/Fluxo principal)
      await criarAba('Encaminhamentos', 'encaminhamentos', [
        { header: 'NOME', key: 'nome', width: 35 },
        { header: 'CROSS', key: 'cross', width: 15 },
        { header: 'DATA NASC', key: 'dataNasc', width: 15 },
        { header: 'TELEFONE', key: 'telefone', width: 20 },
      
        { header: 'EXAME / PROCEDIMENTO', key: 'especialidade', width: 30 },
      
        { header: 'STATUS', key: 'status', width: 25 },
      
        { header: 'DATA REGISTRO', key: 'dataRegistro', width: 15 },
        { header: 'DATA CHEGADA', key: 'dataChegada', width: 15 },
      
        { header: 'LOCAL CONSULTA', key: 'local', width: 30 },
        { header: 'DATA CONSULTA', key: 'dataConsulta', width: 15 },
        { header: 'HORA CONSULTA', key: 'horaConsulta', width: 15 },
      
        { header: 'OBSERVAÇÃO', key: 'obs', width: 45 },
      
        { header: 'MOTIVO CORREÇÃO', key: 'motivoCorrecao', width: 45 },
        { header: 'DATA RETORNO REGULAÇÃO', key: 'dataRetornoRegulacao', width: 20 }
      ]);
  
      // 3. ABA: EXAMES (Versão 3)
      await criarAba('Exames', 'exames', [
        { header: 'CROSS', key: 'cross', width: 15 },
        { header: 'EXAME', key: 'exame', width: 25 },
        { header: 'STATUS', key: 'status', width: 15 },
        { header: 'DATA CHEGADA', key: 'dataChegada', width: 15 }
      ]);
  
      await criarAba('Remessas', 'remessas', [
        { header: 'Nº REMESSA', key: 'numeroRemessa', width: 15 },
        { header: 'DESTINO', key: 'destino', width: 30 },
        { header: 'DE (ORIGEM)', key: 'de', width: 25 },
        { header: 'A/C', key: 'ac', width: 25 },
      
        { header: 'ASSUNTO', key: 'assunto', width: 40 },
        { header: 'DESCRIÇÃO', key: 'descricao', width: 50 },
      
        { header: 'STATUS', key: 'status', width: 18 },
      
        { header: 'DATA SAÍDA', key: 'dataSaida', width: 15 },
        { header: 'DATA RECEBIDO', key: 'dataRecebido', width: 18 },
        { header: 'RECEBIDO POR', key: 'recebidoPor', width: 35 },
      ]);
  
      // 5. ABA: SOS
      await criarAba('Manutenção SOS', 'sos', [
        { header: 'NÚMERO OS', key: 'numeroOS', width: 15 },
        { header: 'UNIDADE', key: 'unidade', width: 20 },
        { header: 'STATUS', key: 'status', width: 15 },
        { header: 'Descrição', key: 'descricaoServico', width: 15 }
      ]);
  
      // Geração do arquivo
      const buffer = await workbook.xlsx.writeBuffer();
      const blob = new Blob([buffer], { type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet' });
      const url = window.URL.createObjectURL(blob);
      const a = document.createElement('a');
      a.href = url;
      a.download = `RELATORIO_COMPLETO_FSPSS_${new Date().toISOString().split('T')[0]}.xlsx`;
      a.click();
      
      setStatusBackup({ tipo: 'sucesso', msg: 'Relatório exportado com sucesso!' });
    } catch (error) {
      console.error(error);
      setStatusBackup({ tipo: 'erro', msg: 'Erro ao gerar Excel completo.' });
    }
  };
  const formatarTelefone = (value: string) => {
    let numbers = value.replace(/\D/g, '');
    if (numbers.length > 11) numbers = numbers.substring(0, 11);
    if (numbers.length > 6) return numbers.replace(/(\d{2})(\d{4,5})(\d{4})/, '($1) $2-$3');
    else if (numbers.length > 2) return numbers.replace(/(\d{2})(\d{0,4})/, '($1) $2');
    else if (numbers.length > 0) return numbers.replace(/(\d{0,2})/, '($1');
    return numbers;
  };

  useEffect(() => {
    const unidadeSalva = localStorage.getItem('fspss_unidade_padrao');
    const telefoneSalvo = localStorage.getItem('fspss_telefone_padrao');
    if (unidadeSalva) setUnidade(unidadeSalva);
    if (telefoneSalvo) setTelefone(telefoneSalvo);
  }, []);

  const handleSalvar = (e: React.FormEvent) => {
    e.preventDefault();
    localStorage.setItem('fspss_unidade_padrao', unidade.trim());
    localStorage.setItem('fspss_telefone_padrao', telefone.trim());
    setSalvo(true);
    setTimeout(() => setSalvo(false), 3000);
  };

  const handleExportarDados = async () => {
    try {
      setStatusBackup(null);
      const tabelas = ['encaminhamentos', 'pedidos', 'materiais', 'remessas', 'sos', 'pacientes', 'exames', 'transferencias'];
      const dadosDexie: Record<string, any[]> = {};
      for (const tabela of tabelas) {
        dadosDexie[tabela] = await (db as any).table(tabela).toArray();
      }
      const dadosConfig = {
        fspss_unidade_padrao: localStorage.getItem('fspss_unidade_padrao') || '',
        fspss_telefone_padrao: localStorage.getItem('fspss_telefone_padrao') || '',
        fspss_mensagem_zap_padrao: localStorage.getItem('fspss_mensagem_zap_padrao') || '',
      };
      const objetoBackup = {
        sistema: 'GestaoClinicaFSPSS',
        dataExportacao: new Date().toLocaleDateString('pt-BR'),
        configuracoes: dadosConfig,
        bancoDados: dadosDexie
      };
      const dataStr = "data:text/json;charset=utf-8," + encodeURIComponent(JSON.stringify(objetoBackup, null, 2));
      const a = document.createElement('a');
      a.setAttribute("href", dataStr);
      a.setAttribute("download", `BACKUP_GERAL_FSPSS_${new Date().toISOString().split('T')[0]}.json`);
      document.body.appendChild(a);
      a.click();
      a.remove();
      setStatusBackup({ tipo: 'sucesso', msg: 'Backup gerado e baixado com sucesso!' });
    } catch (error) {
      setStatusBackup({ tipo: 'erro', msg: 'Falha ao gerar arquivo de exportação.' });
    }
  };

  const handleImportarDados = (e: React.ChangeEvent<HTMLInputElement>) => {
    const arquivos = e.target.files;
    if (!arquivos || arquivos.length === 0) return;
    const leitor = new FileReader();
    leitor.onload = async (evento) => {
      try {
        const conteudo = evento.target?.result as string;
        const dadosImportados = JSON.parse(conteudo);
        if (dadosImportados.sistema !== 'GestaoClinicaFSPSS') throw new Error('Arquivo inválido.');
        if (!confirm('Atenção: A importação irá mesclar ou substituir registros locais. Deseja continuar?')) return;
        if (dadosImportados.configuracoes) {
          Object.entries(dadosImportados.configuracoes).forEach(([chave, valor]) => {
            localStorage.setItem(chave, valor as string);
          });
          setUnidade(dadosImportados.configuracoes.fspss_unidade_padrao || '');
          setTelefone(dadosImportados.configuracoes.fspss_telefone_padrao || '');
        }
        if (dadosImportados.bancoDados) {
          for (const [nomeTabela, registros] of Object.entries(dadosImportados.bancoDados)) {
            if (Array.isArray(registros)) {
              await (db as any).table(nomeTabela).clear();
              await (db as any).table(nomeTabela).bulkAdd(registros);
            }
          }
        }
        setStatusBackup({ tipo: 'sucesso', msg: 'Dados importados! Recarregando...' });
        setTimeout(() => window.location.reload(), 2000);
      } catch (error: any) {
        setStatusBackup({ tipo: 'erro', msg: error?.message || 'Erro ao processar o backup.' });
      }
    };
    leitor.readAsText(arquivos[0]);
  };

  return (
    <main className="min-h-screen bg-gray-50 p-6 sm:p-8 text-gray-950">
      <div className="max-w-xl mx-auto space-y-6">

        <div className="border-b border-gray-200 pb-4">
          <span className="text-[10px] font-black text-blue-600 uppercase tracking-widest block mb-1">
            Preferências do Sistema
          </span>
          <h1 className="text-2xl font-black text-gray-900 tracking-tight uppercase">
            Configurações Gerais
          </h1>
          <p className="text-xs text-gray-500 mt-1">
            Defina os dados padrão para preenchimento automático de formulários e gerencie a migração ou cópia de segurança dos dados locais.
          </p>
        </div>

        {salvo && (
          <div className="p-3 bg-emerald-50 border border-emerald-200 text-emerald-800 rounded-lg flex items-center gap-2 text-xs font-bold uppercase tracking-wide">
            <CheckCircle2 size={16} className="text-emerald-600" />
            Configurações padrão salvas com sucesso!
          </div>
        )}

        {statusBackup && (
          <div className={`p-3 border rounded-lg flex items-center gap-2 text-xs font-bold uppercase tracking-wide ${
            statusBackup.tipo === 'sucesso' ? 'bg-blue-50 border-blue-200 text-blue-800' : 'bg-red-50 border-red-200 text-red-800'
          }`}>
            <RefreshCw size={16} className={statusBackup.tipo === 'sucesso' ? 'animate-spin text-blue-600' : 'text-red-600'} />
            {statusBackup.msg}
          </div>
        )}

        <div className="bg-white border border-gray-200 rounded-xl p-6 shadow-sm space-y-4">
          <div>
            <span className="text-[10px] font-black text-blue-600 uppercase tracking-widest block mb-1">Migração e Cópia Local</span>
            <h3 className="font-black text-sm text-gray-900 uppercase tracking-tight">Transferência entre Computadores</h3>
            <p className="text-xs text-gray-500 mt-0.5">Salve um arquivo com tudo que está salvo neste PC para abrir em outro navegador.</p>
          </div>
          <div className="grid grid-cols-1 sm:grid-cols-2 gap-3">
            <button type="button" onClick={handleExportarDados}
              className="flex items-center justify-center gap-2 bg-gray-100 hover:bg-gray-200 text-gray-800 border border-gray-300 rounded-lg py-2.5 px-4 text-xs font-black uppercase tracking-wider transition-colors cursor-pointer">
              <Download size={14} className="text-blue-600" /> Exportar Dados (.json)
            </button>
            <button type="button" onClick={() => fileInputRef.current?.click()}
              className="flex items-center justify-center gap-2 bg-gray-100 hover:bg-gray-200 text-gray-800 border border-gray-300 rounded-lg py-2.5 px-4 text-xs font-black uppercase tracking-wider transition-colors cursor-pointer">
              <Upload size={14} className="text-emerald-600" /> Importar Dados
            </button>
            <button type="button" onClick={handleExportarExcel}
              className="flex items-center justify-center gap-2 bg-emerald-600 hover:bg-emerald-700 text-white rounded-lg py-2.5 px-4 text-xs font-black uppercase tracking-wider transition-colors col-span-1 sm:col-span-2 cursor-pointer">
              <Download size={14} /> Exportar p/ Excel (.xlsx)
            </button>
            <input type="file" ref={fileInputRef} onChange={handleImportarDados} accept=".json" className="hidden" />
          </div>
        </div>

        <div className="bg-white border border-gray-200 rounded-xl p-6 shadow-sm">
          <form onSubmit={handleSalvar} className="space-y-4">

            <div>
              <label className="block text-[10px] font-black text-gray-400 uppercase tracking-wider mb-2 flex items-center gap-1">
                <Building2 size={12} /> Unidade de Saúde Padrão
              </label>
              <input type="text" placeholder="Ex: USF Barra do Sahy" value={unidade}
                onChange={(e) => setUnidade(e.target.value.toUpperCase())}
                className="w-full px-3 py-2 text-sm bg-gray-50 border border-gray-300 rounded-lg focus:outline-none focus:ring-2 focus:ring-blue-500 focus:bg-white text-gray-900 font-medium uppercase" />
              <span className="text-[10px] text-gray-400 mt-1 block">
                Essa unidade virá pré-selecionada ao gerar remessas, pedidos e chamados S.O.S.
              </span>
            </div>

            <div>
              <label className="block text-[10px] font-black text-gray-400 uppercase tracking-wider mb-2 flex items-center gap-1">
                <Phone size={12} /> Telefone / Ramal da Unidade
              </label>
              <input type="text" placeholder="Ex: (12) 3865-0000" value={telefone}
                onChange={(e) => setTelefone(formatarTelefone(e.target.value))}
                className="w-full px-3 py-2 text-sm bg-gray-50 border border-gray-300 rounded-lg focus:outline-none focus:ring-2 focus:ring-blue-500 focus:bg-white text-gray-900 font-medium" />
              <span className="text-[10px] text-gray-400 mt-1 block">
                Número que aparecerá no rodapé impresso da Ordem de Serviço e documentos da unidade.
              </span>
            </div>

            <div className="pt-2 border-t border-gray-100 mt-6">
              <button type="submit"
                className="w-full bg-blue-600 hover:bg-blue-700 text-white font-bold text-xs uppercase tracking-wider py-3 px-4 rounded-lg flex items-center justify-center gap-2 transition-colors shadow-sm cursor-pointer">
                <Save size={14} /> Salvar Definições Padrão
              </button>
            </div>

          </form>
        </div>

      </div>
    </main>
  );
}
