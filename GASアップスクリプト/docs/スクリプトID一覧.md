# Apps Script スクリプトID一覧（clasp設定用）

`.clasp.json` はgit同期されない（`.gitignore`対象）ので、**別のPCで clasp push したいときは、そのプロジェクトフォルダにこの一覧を見て作る**。

作り方：プロジェクトフォルダ直下に `.clasp.json` を作成して、以下の1行を入れる（IDは下の表から）。

```json
{ "scriptId": "ここに該当プロジェクトのID" }
```

反映コマンド：フォルダに移動して `npx @google/clasp push -f`
（初回は `npx @google/clasp login` でGoogleアカウント認証が必要）

| プロジェクト | スクリプトID |
|---|---|
| Project_01 | 15o-6H_Hi5Ttx_XlE_A9HfGsjLdH5sbE2hXH7zzxtqkvEkj1DKvqCz0m4 |
| Project_02 | 1fSHlgaw0lEXFCLGjvy_1fElpumXuyBWBZtxYtYkKV9Mk-D7ChzRDF1a7 |
| Project_03 | 1lzvOpF61JM2LZcvuP-k62xQA9rkq4oXYvbusjDzrC6xqfB5xj2QjdWDx |
| Project_04 | 1A8njW3CmqRsibmVG3qB9mfQOeyTn5SrJN8e6ngoDU6VSkqDiQI9eFpnZ |
| Project_05 | 17CMxCQ_pG044Ggcd1DirNQC_PvUcZwmuqpKcEhwVCK79IcTgqpULFotA |
| Project_06 | 1qJGkGNleplh0PocO73UC28_g7BkwBJ4MIXVzbopX5Ew6UrZzYP-dn5Is |
| Project_07 | 16fN-9iQNXNoCjjBupA24VcvslUlC9OPmy0Dhg78W7uHvylWVc-5BtUoJ |
| Project_08 | 17x6e7MdvMh6V3DZ_kQqIn13-r_UHJ9IzLYvTYFb9LX3H3sa3ygpvt_tl |
| Project_09 | 1L6wBamaxcIQNXP6enNr8H3_4V7AeYvBCtT-IM--FQ6AmYPBmJQXynOPf |
| Project_10 | 1HcFqCh1EOctaU4wDRcZQZexufNBgkLPSHt03rWB2cJV0Zfa8q-_CEx-Q |
| Project_11 | 1_XN3sWn15-AJbeIdaY7fo7xhIB6OVNYsDa21pDZlPu0h0GcEAIbOMq4r |
| Project_12 | 1G9E5q8n8CH8u0kHicfQrPDUSX-NWZvDg4ywSWuXg77c0GcsBEBWjwafw |
| Project_13 | 15spz2cVP15P6P60caglwOynOC2uIaskqLvnThV1cPIr2kXPdqyNB1vOm |
| Project_14 | 1GdbWZfuCZQFXbNDJFqRKrQe-YiwpB-c2OesyToMWS30vVmBnju7En0PL |
| Project_15 | 1aeNrbz3HulMrE9h3aQD2CNOQSHjN-lyv2aAkFnXujrfEh9tWaIlIudCw |
| Project_16 | 1whZqOo-ztjAEhDcvAkOmQFTtH5J2avrQBk6iQawdZN2j-gduqD1KG6Ay |
| Project_17 | 18KCGO0ccu_XrQd5skoCltAEJTCrsLSXU-nOHIJdpHE8OHpHbOg_aVWuH |
| Project_18 | 1tBq8ciT1bh-1dPRi1zrk2QVd9WkpXzrtBwWRrendz8kOLhlepkuG77ro |
| Project_19（発注共有ファイル / 発注EMSリスト） | 17N3ueJKAdvmsjAohaQzQCHUitW9RxHZpLlAm2zrzDJzLV811d8cBuHqT |
| Project_20 | （clasp未設定） |
| Project_21 | （clasp未設定） |
| Project_22 | （clasp未設定） |
| Project_23 | （clasp未設定） |
| Project_24（引当ファイル / GoQ受注→EMS引き当て） | 1UymlnzhWq6bkNwRPtVlC0aI44LSG4OtVypwayKDWVy1zZZcyIvPC4UF2 |

※Project_20〜23は現時点でclasp未設定。使う場合はApps Scriptエディタの「プロジェクトの設定→スクリプトID」からIDを取得してここに追記する。
