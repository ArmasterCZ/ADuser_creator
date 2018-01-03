ADuser creator
Tento program slouží k zakládání nových uživatelů v Active Directory pomocí integrovaných powershellových skriptů. 
Jeho hlavním účelem je usnadnit zakládání nových uživatelů na základě určité firemní politiky. 

Program se skládá ze sloupce kam je možné začít zapisovat údaje pro nového uživatele a zamknutého sloupce pro porovnání či klonování atributů od vyhledaného uživatele. V sloupci pro nového uživatele funguje systém automatického doplňování. K tomu je využita kombinace dvou možností. První kompletuje jméno, sAMAccountname, email. Další načítá data z excelovské tabulky a doplňuje je na základě klíčové kolonky Kancelář. Poslední převádí číslo karty do správného formátu.

Mezi jeho hlavní funkce patří:
- vytváření uživatelů v AD až s 20 atributy a přiřazením do skupin.
- vyhledání uživatelů v AD a vypsání jejich dat do tabulky.
- klonování dat od vyhledaných uživatelů
- načítání dat z excelovské tabulky
- přesouvání uživatele do jiného kontejneru v AD
- vylepšená funkce klávesových zkratek v tabulce
	- ctrl+Q načíst data z řádky v excelu
	- ctrl+S zapsat uživatele do AD
	- ctrl+V vložit data do vybraných kolonek (při kopírování z excelu)
	- delete smazat data z vybraných kolonek


.NET Framework 4.5.2

Armaster 2016-2017
