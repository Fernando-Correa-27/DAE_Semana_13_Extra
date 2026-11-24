# Análisis del laboratorio S13 — El carrito revisado por pruebas

Documento de trabajo con las respuestas conceptuales de las fases A a E. El entregable final
(`entregable/DAE_S13_Extra_Entregable.docx`) se construye a partir de este material y de las salidas
reales guardadas en `docs/evidencia/`.

## Fase A — Entender el proyecto

### 1. Las reglas de la tienda, con mis palabras

- **Dinero.** Todo importe se representa con `Decimal`, nunca con `float`, porque un `float` guarda
  aproximaciones binarias y pierde céntimos. El redondeo es al céntimo con `ROUND_HALF_UP`: cuando la
  tercera cifra decimal es 5 o más, sube (0,845 → 0,85; 452,205 → 452,21). El helper `money()` de
  `shop/domain.py` concentra esa regla.
- **Precio.** El precio unitario lo decide el catálogo. Una petición HTTP solo puede pedir `sku` y
  `qty`; si envía `price`, se ignora. El carrito lee el precio del `Product` del catálogo al construir
  cada `Line`.
- **Cantidad.** Debe ser un entero de 1 en adelante. Se rechazan `0`, negativos, decimales, cadenas,
  `None` y booleanos (en Python `True` es un `int`, por eso se comprueba aparte). Al quitar más
  unidades de las que hay, la línea queda en cero y se borra; jamás queda una cantidad negativa.
- **Existencias.** La suma de lo que ya está en el carrito más lo que se agrega no puede superar el
  stock del catálogo. Si se intenta, se lanza `OutOfStock` y el carrito no cambia. Al pagar se
  vuelven a comprobar **todas** las líneas antes de restar ninguna; el stock nunca baja de cero.
- **Cupón.** Un solo cupón por carrito. `BIENVENIDA10` descuenta el 10 % del subtotal; `MENOS20`
  descuenta S/ 20 sin pasar del subtotal. Aplicar otro cupón, o el mismo otra vez, **reemplaza** al
  anterior: no se acumula. El código se normaliza con `strip().upper()`.
- **Totales.** `subtotal` = suma de precio × cantidad; `descuento` según el cupón vigente; `igv` = 18 %
  de (subtotal − descuento); `total` = (subtotal − descuento) + igv. Cada parte se redondea al céntimo
  con `money()`.

### 2. Por qué pasan las pruebas de la auditoría con el carrito sin implementar

La suite inicial tiene 67 pruebas: 7 pasan y 60 fallan. Las 7 que pasan son las de
`shop/tests/test_audit.py` que **no usan `Cart`**: `audit_cart()` trabaja sobre el JSON declarado y
recalcula los totales con `reference_totals()` de `shop/domain.py`, una función independiente de toda
implementación. Por eso puede detectar `PRICE_MISMATCH`, `OVER_STOCK`, `BAD_QTY`, `MULTIPLE_COUPONS`
y `TOTAL_MISMATCH` aunque `shop/cart_team.py` sea un esqueleto. Las dos pruebas de replay sí fallan,
porque `--replay` sí construye `impl.Cart` y recibe `NotImplementedError` → `REPLAY_UNAVAILABLE`.
Evidencia: `docs/evidencia/02_test_inicial_equipo.txt`.

### 3. Por qué el carrito usa un catálogo de datos puros y no el modelo de Django

`shop/domain.py` no importa Django: define `Product` como un `dataclass` con `sku`, `name`, `price` y
`stock`. Así la lógica del carrito es pura: se puede probar con `catalog_data()` (los ocho productos
fijos) sin base de datos, sin migraciones y sin servidor; además el `checkout` puede modificar el stock
de una copia del catálogo y la prueba observa esa copia. La frontera con la base de datos vive en
`shop/catalog.py` (`catalog_from_db()` y `save_stock()`), que traduce filas del modelo `Product` al
`ProductData` que el carrito entiende.

### 4. Qué guarda `_qty` y por qué no guarda el precio

`_qty` es un diccionario `sku -> cantidad` entera, en orden de inserción. No guarda el precio porque el
precio pertenece al catálogo: si se copiara en el carrito podría quedar desactualizado o ser
manipulado. `lines()` construye cada `Line` leyendo `self._catalog[sku].price` en ese momento; el
carrito solo recuerda **qué** y **cuánto**, no **a qué precio**.

### 5. Qué hace `shop/impl.py` con `SHOP_MODE`

Lee la variable de entorno `SHOP_MODE` al importarse (por defecto `team`), valida que sea una de
`team`, `ai` o `solutions`, e importa el módulo correspondiente (`shop.cart_team`, `shop.cart_ai` o
`teacher.solutions.cart`). Expone `cart` para que las pruebas, la API, la auditoría y la compuerta usen
la misma implementación. Si el modo no existe o el módulo no está en esa copia, lanza `RuntimeError`
con un mensaje claro.

## Fase B — El carrito del equipo

### Por qué `remove` borra la línea al llegar a cero

Una línea con cantidad cero no representa nada comprable y, si se quedara, aparecería en `lines()` y en
la API con `qty: 0`, ensuciando la interfaz y el estado guardado en sesión. Al borrarla, el diccionario
`_qty` solo contiene productos realmente presentes y `summary()` no muestra líneas fantasma. Si en vez
de borrarse se permitiera una cantidad negativa, el total restaría dinero: por ejemplo, un carrito con
`MOU-01` a −3 sumaría −119,70 al subtotal, el IGV bajaría y el cliente podría “ganar” saldo; además el
stock se repondría al pagar y la línea no tendría sentido. La prueba
`test_remove_more_than_there_is_never_leaves_a_negative_quantity` exige cantidad 0, lista vacía y total
0,00.

### Por qué `checkout` vuelve a comprobar las existencias

Entre `add` y `checkout` puede pasar tiempo y otro cliente puede comprar las mismas unidades. `add` solo
garantiza el stock en el instante de agregar; `checkout` es el último punto antes de cobrar y debe
verificar **todas** las líneas antes de restar **ninguna**, para no dejar un descuento de stock a medias.
Por eso el bucle de comprobación termina y recién después otro bucle resta. En la API, `cart_checkout`
envuelve la operación en `transaction.atomic()` y llama a `save_stock(catalog)`; sin embargo, la
mitigación completa contra dos compradores simultáneos requiere bloquear las filas del catálogo con
`select_for_update()` (o un `UPDATE ... WHERE stock >= qty` atómico). Sin ese bloqueo, dos procesos
pueden leer el mismo stock y ambos pasar la comprobación: es una condición de carrera. El laboratorio
usa SQLite y una sola petición por prueba, así que no la reproduce, pero en producción la transacción
con bloqueo de fila es la solución correcta.

### Por qué el doble clic en «Aplicar cupón» no debe acumular

`apply_coupon` guarda el código en un único atributo `self._coupon`: el segundo clic reemplaza el valor
anterior por el mismo, y `discount()` calcula el descuento una sola vez a partir de ese valor. El
resultado es idempotente: dos clics dejan el mismo total que uno. Si el carrito guardara una lista y
sumara el descuento de cada elemento, un doble clic (o un reintento de red) descontaría dos veces y la
tienda perdería dinero. Las pruebas `test_applying_the_same_coupon_twice_discounts_once` y
`test_a_second_coupon_replaces_the_first` fijan esa regla.

### Por qué el precio no puede venir del cliente

El cuerpo de `POST /api/cart/items/` es dato no confiable. `add_from_payload` solo acepta un diccionario
con `sku` de tipo `str` y pasa `sku`/`qty` a `cart.add`; cualquier `price` del cuerpo se ignora. El
precio se toma de `self._catalog[sku].price` dentro de `lines()`. Si se aceptara el precio del cliente,
bastaría enviar `{"sku": "LAP-01", "price": 0.01}` para comprar una laptop de S/ 2 499,90 por un céntimo;
la prueba `test_the_price_comes_from_the_catalog_not_from_the_client` y el carrito
`02-precio-del-cliente.json` documentan ese ataque.

### Llamadas reales a la API (paso 9)

Con `runserver` en `127.0.0.1:8013` y una misma sesión de cookies (`tools/api_smoke.ps1`):

1. `POST /api/cart/items/ {"sku":"MOU-01","qty":2}` → 201; subtotal `79.80`, IGV `14.36`, total
   `94.16`.
2. `POST /api/cart/coupon/ {"code":"BIENVENIDA10"}` → descuento `7.98`, IGV `12.93`, total `84.75`.
3. `POST /api/cart/checkout/ {}` → 201; devuelve el resumen de la orden con total `84.75` y deja el
   carrito vacío (`GET /api/cart/` → `total: "0.00"`). El servidor se detuvo al terminar y
   `seed_shop` restauró el catálogo (8 productos).

Salida completa: `docs/evidencia/19_api_smoke.txt` y `20_seed_restore.txt`.

### 5 bis. Float frente a Decimal (paso 5)

Salida real de `python tools/decimal_check.py` (`docs/evidencia/09_turno3_decimales.txt`):

```text
0.1 + 0.2 = 0.30000000000000004
Decimal('0.1') + Decimal('0.2') = 0.3
round(0.845, 2) = 0.84
Decimal('0.845').quantize(Decimal('0.01'), ROUND_HALF_UP) = 0.85
```

Un `float` guarda el número en binario con 53 bits de mantisa: 0,1 y 0,2 no tienen representación
binaria finita, así que se guardan como la fracción más cercana y su suma arrastra el error
`0.30000000000000004`. Tampoco 0,845 es exacto en binario: el valor almacenado queda ligeramente por
debajo de 0,845, y `round()` de Python, además, redondea al par (banker's rounding), por eso devuelve
0,84. `Decimal` guarda dígitos decimales exactos y `quantize(..., ROUND_HALF_UP)` aplica la regla de
la tienda: 0,845 → 0,85. Por eso `money()` usa `Decimal` y `ROUND_HALF_UP` en todo el carrito.

**Cálculo a mano del mousepad con `BIENVENIDA10`** (PAD-01, precio 8,45, cantidad 1):

1. Subtotal = 8,45 × 1 = **8,45**.
2. Descuento = 10 % de 8,45 = 0,845 → redondeo half-up → **0,85**.
3. Base = 8,45 − 0,85 = **7,60**.
4. IGV = 18 % de 7,60 = 1,368 → half-up → **1,37**.
5. Total = 7,60 + 1,37 = **8,97**.

Coincide con `test_the_discount_rounds_half_up`, que exige `discount() == 0.85` y `total() == 8.97`.

## Fase C — El duelo

Salida real de `python manage.py duelo` (`docs/evidencia/21_duelo.txt`):

| Modo | Implementación | Pasan | Fallan |
|---|---|---|---|
| `team` | tu carrito | 67 | 0 |
| `ai` | el carrito de la IA | 49 | 18 |

Fallos de la IA agrupados por archivo (`docs/evidencia/22_duelo_ai.txt`):

- `shop/tests/test_api.py` (6): `test_checkout_when_the_stock_changed_is_409_and_keeps_the_cart`
  (ERROR, `CHECK constraint failed: stock`), `test_a_negative_price_from_the_client_changes_nothing_either`,
  `test_more_than_the_stock_is_409`, `test_remove_more_than_there_is_leaves_no_negative_line`,
  `test_the_coupon_is_applied_once_even_if_the_client_repeats_it`,
  `test_the_price_comes_from_the_catalog_not_from_the_client`.
- `shop/tests/test_audit.py` (1): `test_replay_of_the_rounding_cart_matches_your_cart`.
- `shop/tests/test_coupons.py` (2): `test_a_second_coupon_replaces_the_first`,
  `test_applying_the_same_coupon_twice_discounts_once`.
- `shop/tests/test_lines.py` (1): `test_remove_more_than_there_is_never_leaves_a_negative_quantity`.
- `shop/tests/test_stock.py` (4): `test_checkout_checks_the_stock_again`,
  `test_the_stock_counts_what_is_already_in_the_cart`, `test_the_stock_is_never_negative_after_any_sequence`,
  `test_two_adds_cannot_get_around_the_stock`.
- `shop/tests/test_totals.py` (4): `test_a_discount_of_one_and_a_half_cents`, `test_igv_rounds_half_up`,
  `test_the_discount_rounds_half_up`, `test_total_is_the_sum_of_its_rounded_parts`.

## Fase D — Los cinco fallos de la IA, clasificados sin tocar el archivo

### Tabla de los cinco fallos

| # | Archivo:línea | Fallo | Prueba que lo detecta | Mensaje / diferencia real | Causa | ¿Se vería a mano en una tienda? | Marca |
|---|---|---|---|---|---|---|---|
| 1 | `shop/cart_ai.py:14`, `17-18`, `83-84`, `92-94`, `104` | Redondeo con flotantes | `test_totals` (4), `test_audit.test_replay_of_the_rounding_cart_matches_your_cart` | `Decimal('0.84') != Decimal('0.85')`; `Decimal('452.2') != Decimal('452.21')`; `Decimal('0.1') != Decimal('0.11')`; `Decimal('71.63') != Decimal('71.64')`; `REPLAY_MISMATCH` | `_money` usa `round(value, 2)` sobre `float`; `_subtotal` suma `float`; `IGV = 0.18` es `float` | Sí, con valores de medio céntimo: un mousepad de S/ 8,45 con 10 % debe descontar 0,85 y la IA descuenta 0,84; el cliente ve un céntimo de diferencia | Cálculo |
| 2 | `shop/cart_ai.py:45`, `53`, `139` | Precio tomado del cliente | `test_api.test_the_price_comes_from_the_catalog_not_from_the_client`, `test_api.test_a_negative_price_from_the_client_changes_nothing_either`; auditoría: `PRICE_MISMATCH` en `02-precio-del-cliente.json` | `'0.01' != '2499.90'`; `'-100.00' != '2499.90'`; la petición con `price` devuelve 201 | `add(..., price=None)` guarda `Decimal(str(price))` en `_prices` y `add_from_payload` pasa `payload.get("price")` | Solo si alguien manipula la petición; en la pantalla normal no se nota, pero permite comprar una laptop de S/ 2 499,90 por S/ 0,01 | **Seguridad** y **solo API** |
| 3 | `shop/cart_ai.py:50` y `124-133` | Existencias que no cuentan lo ya agregado (y checkout no re-verifica) | `test_stock` (4), `test_api.test_more_than_the_stock_is_409`, `test_api.test_checkout_when_the_stock_changed_is_409_and_keeps_the_cart`; `OVER_STOCK` en `03-stock.json` | `OutOfStock not raised`; `201 != 409`; `-2 not greater than or equal to 0`; `CHECK constraint failed: stock` | `add` compara solo `qty > product.stock`, no `self._qty.get(sku,0)+qty`; `checkout` resta sin volver a comprobar | Sí: con 2 mousepads en stock, agregar 2 y luego 1 deja 3 en el carrito; al pagar se vende stock inexistente y el inventario queda negativo | Existencias |
| 4 | `shop/cart_ai.py:55-66` (línea 63) | Cantidad negativa al quitar | `test_lines.test_remove_more_than_there_is_never_leaves_a_negative_quantity`, `test_api.test_remove_more_than_there_is_leaves_no_negative_line` | `-3 != 0`; la línea queda con `qty: -2` y el total baja | `self._qty[sku] -= qty` sin tope; solo borra si queda exactamente 0 | Sí: quitar 5 unidades cuando hay 2 deja la línea en −3 y el total se vuelve negativo; el cliente “gana” dinero | Cantidad |
| 5 | `shop/cart_ai.py:77-81`, `89-95` | Descuento acumulado dos veces | `test_coupons` (2), `test_api.test_the_coupon_is_applied_once_even_if_the_client_repeats_it`; `MULTIPLE_COUPONS` en `05-cupon-doble.json` | `Decimal('37.8') != Decimal('18.90')`; `'37.80' != '18.90'`; `Decimal('38.9') != Decimal('20.00')`; `REPLAY_MISMATCH` | `apply_coupon` hace `append` a una lista y `_discount` suma el descuento de todos los cupones | Sí: un doble clic en «Aplicar cupón» descuenta dos veces el 10 %; el cliente paga menos de lo que corresponde | Idempotencia |

El fallo de **seguridad** es el 2: acepta el precio que manda el cliente. Es también el único que
**solo se ve con la API**, porque el camino `add_from_payload` es el que traduce el cuerpo HTTP; las
pruebas puras del carrito no envían `price`. Los fallos 1, 3, 4 y 5 se reproducen directamente sobre
`Cart` y por eso aparecen en `test_totals`, `test_stock`, `test_lines` y `test_coupons`. El fallo 3
tiene además manifestaciones HTTP (`409` que no llega y el `CHECK constraint failed: stock`), pero su
raíz es la misma regla de existencias.

### Las cuatro preguntas de revisión aplicadas a `cart_ai.py`

1. **¿Todo el dinero es `Decimal` de principio a fin, con `ROUND_HALF_UP`? No.** La línea 14 define
   `IGV = 0.18` como `float`; la 18 hace `Decimal(str(round(value, 2)))`, que primero redondea en
   binario; la 84 suma `float(l.unit_price) * l.qty`; las líneas 92 y 94 acumulan `float`; la 104
   mezcla ese `float` con el IGV. El resultado es 0,84 donde la regla pide 0,85.
2. **¿El servidor es la única fuente del precio? No.** La línea 45 acepta `price=None`; la 53 guarda
   `Decimal(str(price))` si el cliente lo envía; la 139 le pasa `payload.get("price")`. El catálogo
   deja de ser la única fuente.
3. **¿Cada regla de cantidad y existencias se cumple con cualquier secuencia? No.** La línea 50
   compara `qty > product.stock` sin sumar `self._qty.get(sku, 0)`, así que dos `add` seguidos se
   saltan el stock. La línea 63 resta sin tope y deja cantidades negativas. Las líneas 124-133 no
   vuelven a comprobar el stock antes de restar en `checkout`.
4. **¿Repetir una acción da el mismo resultado? No.** La línea 81 hace `self._coupons.append(code)` y
   las líneas 89-95 recorren y suman **todos** los cupones; aplicar dos veces `BIENVENIDA10` descuenta
   37,80 en vez de 18,90.

### Auditoría antes de los arreglos (IA rota)

Salida completa: `docs/evidencia/23_audit_ai_antes.txt` (sin replay) y
`24_audit_ai_replay_antes.txt` (con replay).

| Carrito | Sin `--replay` | Con `--replay` | Fallo |
|---|---|---|---|
| `01-correcto.json` | Sin hallazgos | Sin hallazgos | — |
| `02-precio-del-cliente.json` | `PRICE_MISMATCH` (0.01 vs 2499.90) + 3 `TOTAL_MISMATCH` | Igual; **no** hay `REPLAY_MISMATCH` | 2 (precio del cliente) |
| `03-stock.json` | `OVER_STOCK` (4 unidades, hay 2) | `REPLAY_REJECTED` (la IA rechaza 4 > 2) | 3 (existencias) |
| `04-cantidad-negativa.json` | `BAD_QTY` (−3) + 3 `TOTAL_MISMATCH` | `REPLAY_REJECTED` (la IA rechaza −3 en `add`) | 4 (cantidad negativa) |
| `05-cupon-doble.json` | `MULTIPLE_COUPONS` + 3 `TOTAL_MISMATCH` | `REPLAY_MISMATCH` en descuento, IGV y total | 5 (cupón acumulado) |
| `06-redondeo.json` | `TOTAL_MISMATCH` (0.84 vs 0.85) | `REPLAY_MISMATCH` (0.84 vs 0.85) | 1 (redondeo) |

**`TOTAL_MISMATCH` frente a `REPLAY_MISMATCH`.** El primero compara lo que el archivo **declara**
(`declared`) con `reference_totals()`; señala que el carrito está mal documentado o que quien lo
generó calculó mal. El segundo compara lo que **calcula la implementación activa** al reconstruir el
carrito (`--replay`) con la misma fórmula; señala un defecto del código, aunque la declaración fuera
correcta. En `06` se ven los dos: la declaración dice 0,84 y el replay de la IA también da 0,84,
mientras la fórmula manda 0,85. En `02` solo hay `TOTAL_MISMATCH`/`PRICE_MISMATCH`: el replay usa
`cart.add(sku, qty)` **sin precio**, así que la IA cae en `product.price` y calcula bien; el defecto
de aceptar el precio del cliente no se reproduce en el replay.

### `--json`, `--strict` y el límite de la auditoría

`25_audit_ai_json_antes.txt` devuelve un objeto con `mode` y `carts[]` (`file`, `ok`, `findings[]`).
Un programa puede leer `findings[].code` y decidir; un agente puede tomar el primer hallazgo, ubicar
la línea y proponer un parche. `26_audit_ai_strict_antes.txt` termina con
`CommandError: 5 carrito(s) con hallazgos.` y **código de salida 1**, útil para un pipeline de CI.

La auditoría **no sustituye las reglas del servidor**: examina un carrito ya declarado o reconstruido
y avisa **después** de que el cálculo ocurrió; no impide que una petición con `price` o con cantidad
negativa llegue a cobrarse. La defensa previa es `add`, `remove`, `checkout` y `add_from_payload` en el
carrito: si la regla está en el servidor, la petición mala se rechaza **antes** de tocar el stock o el
dinero. La auditoría es una red de seguridad para revisar código y datos, no la puerta de entrada.

### Compuerta de la IA antes de los arreglos

`27_gate_ai_antes.txt`: 49 pasan / 18 fallan, 4 de 6 carritos como se esperaba, veredicto `RECHAZAR`
(código 1). `28_gate_ai_json_antes.txt` permite a un agente leer:

- `verdict`: `"rechazar"`.
- `tests.failing`: los 18 nombres dotted de las pruebas que fallan.
- `carts[].result`: `05-cupon-doble.json` y `06-redondeo.json` aparecen como `calcula_distinto`;
  los demás como `ok` (el replay no reproduce el precio del cliente ni las secuencias de stock).

## Fase E — Estado original y arreglos de la IA

Commit del estado original: `6978401` (`chore: snapshot original AI cart before fixes`), etiqueta
`ai-original`. Antes de tocar el archivo: `SHOP_MODE=ai python manage.py test` →
`FAILED (failures=17, errors=1)` = 18 pruebas únicas en rojo (`docs/evidencia/29_test_ai_antes.txt`).

| Arreglo | Commit | Fallas antes | Fallas después | Qué cambió |
|---|---|---|---|---|
| 1. Redondeo con flotantes | `bbb4251` | 18 | 13 | `IGV` a `Decimal("0.18")`; `_money` usa `money`; `_subtotal` y `_discount` en `Decimal`, con el descuento redondeado antes de calcular la base |
| 2. Precio del cliente | `2138c41` | 13 | 11 | `add` sin `price`; se elimina `_prices`; `lines` lee el catálogo; `add_from_payload` no pasa `price` |
| 3. Existencias | `e69e6db` | 11 | 5 | `add` compara `lo_que_hay + qty`; `checkout` re-verifica todas las líneas antes de restar |
| 4. Cantidad negativa | `32489e4` | 5 | 3 | `remove` calcula el resto y borra la línea si es `<= 0` |
| 5. Descuento acumulado | `e2f1be2` | 3 | 0 | `apply_coupon` reemplaza (`_coupons = [code]`); `_discount` usa solo el primer cupón; `from_state` conserva el último |

Después: `SHOP_MODE=ai python manage.py test` → 67/67 (`38_test_ai_fix5_cupon.txt`); `gate --modo ai`
→ `ACEPTAR`, código 0, 6/6 carritos (`39_gate_ai_despues.txt`, `40_gate_ai_json_despues.txt`).

### Por qué una implementación correcta debe rechazar los carritos 03 y 04

`shop/gate.py` define `MUST_REJECT = {"BAD_QTY", "UNKNOWN_SKU", "OVER_STOCK"}`. El carrito 03 declara
`PAD-01` con 4 unidades cuando el catálogo tiene 2 → `OVER_STOCK`; el 04 declara `MOU-01` con −3 →
`BAD_QTY`. Para esos dos, `expected_outcome()` devuelve `"rechazar"` y `judge_cart()` solo da `"ok"` si
el replay produce `REPLAY_REJECTED`. Una implementación correcta lanza `OutOfStock` al agregar 4 sobre
2 y `InvalidQuantity` al agregar −3, así que el replay las rechaza. Si las aceptara, el resultado sería
`"acepta_lo_que_debia_rechazar"` y la compuerta no podría aceptar. Rechazar esos carritos **no** es un
fallo: es la prueba de que las reglas de cantidad y stock están puestas.

### Por qué el replay no detecta el precio del cliente

`shop/audit.py::_replay` reconstruye con `cart.add(item["sku"], item["qty"])`, **sin** pasar
`unit_price`. El carrito 02 declara `LAP-01` a 0,01, pero el replay llama a `add("LAP-01", 1)` y la IA
usa el precio del catálogo; por eso no hay `REPLAY_MISMATCH` y el replay “pasa”. La parte que sí lo
detecta es el **escaneo de la declaración**: `audit_cart()` compara `item["unit_price"]` con
`product.price` y emite `PRICE_MISMATCH`. Sin embargo, `gate.judge_cart()` no rechaza por
`PRICE_MISMATCH` (no está en `MUST_REJECT`), así que la compuerta lo detecta por el **otro oráculo**:
la prueba `test_the_price_comes_from_the_catalog_not_from_the_client` falla mientras `add` acepte
`price`. En resumen: el replay audita el cálculo, no la procedencia del precio; la prueba de la API y
el `PRICE_MISMATCH` de la auditoría auditan la procedencia.

### Auditoría después de los arreglos

`41_audit_ai_replay_despues.txt`: 01 limpio; 02 conserva `PRICE_MISMATCH` y sus `TOTAL_MISMATCH`
(la declaración sigue siendo falsa); 03 y 04 dan `REPLAY_REJECTED` (correcto); 05 y 06 ya **no**
tienen `REPLAY_MISMATCH`, aunque sus declaraciones siguen mal. `42_audit_ai_strict_despues.txt`
termina con código 1 porque los datos declarados de 02–06 son incorrectos: una implementación
correcta no convierte en correctos los carritos mal declarados, solo deja de calcularlos mal.

## Fase F — Rechazo documentado (de ejercicio)

Con la IA ya en verde, se probó un parche **tentador**: dejar que `_subtotal` sume con `float` y
“arreglar” el redondeo convirtiendo a `Decimal` con `Decimal(str(value))`. Las 67 pruebas pasan y la
compuerta acepta, pero el diff viola la pregunta 1: el dinero pasa por `float` y ya no se usa el
helper `money`. Es un **rechazo de ejercicio**, no un rechazo espontáneo: se aplicó a propósito para
demostrar que la compuerta no sustituye la revisión humana.

Diff guardado en `docs/evidencia/rechazo_01.diff`:

```diff
-from decimal import Decimal
+from decimal import ROUND_HALF_UP, Decimal
...
 def _money(value):
-    return money(value)
+    return Decimal(str(value)).quantize(Decimal("0.01"), rounding=ROUND_HALF_UP)
...
     def _subtotal(self):
-        return money(sum((l.unit_price * l.qty for l in self.lines()), Decimal("0")))
+        return _money(sum((float(l.unit_price) * l.qty for l in self.lines()), 0.0))
```

Resultado de la compuerta con el parche: `67/67`, `6/6` carritos, `ACEPTAR` (código 0)
(`44_rechazo_01_gate.txt`). La pregunta 1 lo rechaza porque `float(l.unit_price) * l.qty` introduce
binario en una suma de dinero y `Decimal(str(value))` convierte el float ya contaminado; el helper
`money` de `shop/domain.py` deja de usarse. Se revirtió con `git checkout -- shop/cart_ai.py` y la
compuerta volvió a `ACEPTAR` (`45_rechazo_01_revertido_gate.txt`).

### 6. Estado inicial medido

| Comando | Resultado | Evidencia |
|---|---|---|
| `python manage.py test` | 67 pruebas, 2 fallos + 67 errores de subTest; 7 pasan | `02_test_inicial_equipo.txt` |
| `python manage.py gate` | `RECHAZAR`, código 1; 7 pasan / 60 fallan; 0/6 carritos | `03_gate_inicial_equipo.txt` |
| `python manage.py gate --modo ai` | `RECHAZAR`, código 1; 49 pasan / 18 fallan; 4/6 carritos | `04_gate_inicial_ai.txt` |
