CREATE TABLE hino (
cod_hino int,
cod_artista int,
titulo text,
texto text,
primary key (cod_hino));

CREATE TABLE artista (
cod_artista int,
artista text,
primary key(cod_artista) );

CREATE TABLE links (
cod_link int,
cod_artista int,
link text,
primary key(cod_link) );

CREATE TABLE historico (
cod_historico int,
cod_hino int,
cod_artista int,
data_apresentado datetime default current_timestamp,
primary key (cod_historico) );