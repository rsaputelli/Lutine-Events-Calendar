[CLIENTS_README.md](https://github.com/user-attachments/files/22413915/CLIENTS_README.md)# Managing Clients in Master Calendar App

The **Client** dropdown in the app comes from the Supabase table
`public.clients`.\
The `load_clients()` function queries this table
(`select name from clients order by name`).

To update the dropdown list, edit the `clients` table directly.

------------------------------------------------------------------------

## View all clients

``` sql
select * from public.clients order by name;
```

------------------------------------------------------------------------

## Insert a new client

``` sql
insert into public.clients (name) values ('HAFP');
```

------------------------------------------------------------------------

## Update (rename) a client

``` sql
update public.clients
set name = 'WOEMA'
where name = 'WOHC';
```

------------------------------------------------------------------------

## Delete a client

``` sql
delete from public.clients
where name = 'Old Client';
```

------------------------------------------------------------------------

## (Optional) Add uniqueness constraint

Prevents duplicate names.

``` sql
alter table public.clients
add constraint clients_name_uk unique (name);
```

------------------------------------------------------------------------

## Upsert (insert if not exists)

Useful if you're automating client adds:

``` sql
insert into public.clients (name)
values ('NHCMA')
on conflict (name) do nothing;
```

------------------------------------------------------------------------

⚠️ **Note:** Changes take effect immediately---there is no restart
needed in the Streamlit app. Just refresh your browser tab and the
updated client list will appear.

