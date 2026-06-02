import Dexie from 'dexie';

export const db = new Dexie('GestaoClinicaFSPSS');

// V1 (mantida para histórico)
db.version(1).stores({
  encaminhamentos: '++id, nome, cross, dataRegistro, status, especialidade',
  pedidos: '++id, numeroPedido, data, categoria, status',
  materiais: '++id, codigo, descricao',
  remessas: '++id, numeroRemessa, destino, dataSaida'
});

// V2 (mantida para histórico)
db.version(2).stores({
  encaminhamentos: '++id, nome, cross, dataRegistro, status, especialidade',
  pedidos: '++id, numeroPedido, data, categoria, status',
  materiais: '++id, codigo, descricao',
  remessas: '++id, numeroRemessa, destino, dataSaida',
  sos: '++id, numeroOS, ano, unidade, dataSolicitacao, status'
});

// V3 (mantida para histórico)
db.version(3).stores({
  encaminhamentos: '++id, nome, cross, dataRegistro, status, especialidade',
  pedidos: '++id, numeroPedido, data, categoria, status',
  materiais: '++id, codigo, descricao',
  remessas: '++id, numeroRemessa, destino, dataSaida',
  sos: '++id, numeroOS, ano, unidade, dataSolicitacao, status',
  pacientes: 'cross, nome, dataNasc, telefone',
  exames: '++id, cross, exame, status, dataRegistro, dataChegada'
});

// V4 (mantida para histórico)
db.version(4).stores({
  encaminhamentos: '++id, nome, cross, dataRegistro, status, especialidade',
  pedidos: '++id, numeroPedido, data, categoria, status',
  materiais: '++id, codigo, descricao',
  remessas: '++id, numeroRemessa, destino, dataSaida',
  sos: '++id, numeroOS, ano, unidade, dataSolicitacao, status',
  pacientes: 'cross, nome, dataNasc, telefone',
  exames: '++id, cross, exame, status, dataRegistro, dataChegada',
  transferencias: '++id, unidadeSolicitante, solicitante, paciente, tipoSolicitacao, dataRegistro, status'
});

// V5 (Nova versão com a tabela de profissionais de produtividade)
db.version(5).stores({
  encaminhamentos: '++id, nome, cross, dataRegistro, status, especialidade',
  pedidos: '++id, numeroPedido, data, categoria, status',
  materiais: '++id, codigo, descricao',
  remessas: '++id, numeroRemessa, destino, dataSaida',
  sos: '++id, numeroOS, ano, unidade, dataSolicitacao, status',
  pacientes: 'cross, nome, dataNasc, telefone',
  exames: '++id, cross, exame, status, dataRegistro, dataChegada',
  transferencias: '++id, unidadeSolicitante, solicitante, paciente, tipoSolicitacao, dataRegistro, status',
  
  // Nova tabela adicionada aqui
  profissionais: '++id, nome, matricula, cargo'
});