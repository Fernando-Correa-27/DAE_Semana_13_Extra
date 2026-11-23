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

### 6. Estado inicial medido

| Comando | Resultado | Evidencia |
|---|---|---|
| `python manage.py test` | 67 pruebas, 2 fallos + 67 errores de subTest; 7 pasan | `02_test_inicial_equipo.txt` |
| `python manage.py gate` | `RECHAZAR`, código 1; 7 pasan / 60 fallan; 0/6 carritos | `03_gate_inicial_equipo.txt` |
| `python manage.py gate --modo ai` | `RECHAZAR`, código 1; 49 pasan / 18 fallan; 4/6 carritos | `04_gate_inicial_ai.txt` |
