CREATE TABLE livro (
abrev text,
nome text,
ordem int,
primary key (abrev));

CREATE TABLE traducao (
cod_traducao text,
nome_traducao text,
ordem int,
ativo int,
primary key(cod_traducao) );

CREATE TABLE versiculo (
abrev text,
cod_traducao text,
capitulo int,
versiculo int,
texto text,
primary key (abrev, cod_traducao, capitulo, versiculo) );