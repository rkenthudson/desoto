create table files
(
    id               integer
        constraint files_pk
            primary key autoincrement,
    file_folder_name text,
    name             text,
    suffix           text,
    size             integer,
    modified         date,
    parts            text
);

