"""
Ejemplos de uso del repositorio CRUD para la migración del ERP.
"""

from datetime import date
from decimal import Decimal

from sqlalchemy.orm import Session

from .db import SessionLocal, get_session
from .models import Cliente, Parametro
from .repository import RepositoryFactory


def ejemplo_cliente_crud():
    """Ejemplo: CRUD de Clientes."""
    db: Session = get_session()

    try:
        # Factory de repositorios
        repos = RepositoryFactory(db)
        cliente_repo = repos.cliente()

        # CREATE
        nuevo_cliente = cliente_repo.create(
            {
                "CODIGO": 1001,
                "NOMB": "EMPRESA XYZ S.A.",
                "DIR": "Av. Corrientes 123",
                "LOC": "Buenos Aires",
                "CP": "1025",
                "PCIA": "Buenos Aires",
                "TEL1": "011-4123-4567",
                "CUIT": "30-12345678-9",
                "CIVA": 1,  # 1 = Inscripto
                "CVTA": "01",
                "CREDIT": Decimal("50000.00"),
                "USUARIO": "ADMIN",
                "FACTUAL": date.today(),
            }
        )
        print(f"✓ Cliente creado: {nuevo_cliente.NOMB} (ID={nuevo_cliente.id})")

        # READ (por CODIGO)
        cliente = cliente_repo.by_codigo(1001)
        if cliente:
            print(f"✓ Cliente recuperado: {cliente.NOMB}, CUIT={cliente.CUIT}")

        # UPDATE
        actualizado = cliente_repo.update(
            nuevo_cliente.id, {"CREDIT": Decimal("75000.00"), "DEUDA": Decimal("0.00")}
        )
        if actualizado:
            print(f"✓ Cliente actualizado: CREDIT={actualizado.CREDIT}")

        # READ ALL (con paginación)
        todos = cliente_repo.read_all(skip=0, limit=10)
        print(f"✓ Total de clientes (primeros 10): {len(todos)}")

    finally:
        db.close()


def ejemplo_parametros():
    """Ejemplo: Gestión de Parámetros globales."""
    db: Session = get_session()

    try:
        repos = RepositoryFactory(db)
        param_repo = repos.parametro()

        # CREATE (si no existe)
        param = param_repo.by_clave("IVAINS")
        if not param:
            param = param_repo.create(
                {
                    "CLAVE": "IVAINS",
                    "NUME1": 21,  # IVA Inscripto 21%
                    "USUARIO": "ADMIN",
                    "FECACT": date.today(),
                }
            )
            print(f"✓ Parámetro creado: IVAINS={param.NUME1}%")
        else:
            print(f"✓ Parámetro existente: IVAINS={param.NUME1}%")

        # READ específico
        ptovta = param_repo.get_punto_venta()
        if ptovta:
            print(f"✓ Punto de Venta configurado: {ptovta.NUME1}")

    finally:
        db.close()


def ejemplo_ctascte():
    """Ejemplo: Análisis de Cuentas Corrientes."""
    db: Session = get_session()

    try:
        repos = RepositoryFactory(db)
        cc_repo = repos.ctascte()
        cliente_repo = repos.cliente()

        # Supongamos que existe cliente con CODIGO=1001
        cliente = cliente_repo.by_codigo(1001)
        if cliente:
            # Calcular deuda (saldo = suma de Ctascte.IMPTE según TIPO)
            deuda = cc_repo.deuda_cliente(cliente.CODIGO)
            print(f"✓ Cliente: {cliente.NOMB} | Deuda: ${deuda['deuda']}")

    finally:
        db.close()


def ejemplo_facturacion():
    """Ejemplo: Consultas de Facturación."""
    db: Session = get_session()

    try:
        repos = RepositoryFactory(db)
        fciva_repo = repos.fciva_vta()

        # Facturas en un rango de fechas
        desde = date(2024, 1, 1)
        hasta = date(2024, 12, 31)
        facturas = fciva_repo.by_fecha_rango(desde, hasta)
        print(f"✓ Facturas en 2024: {len(facturas)}")

        # Factura específica
        if facturas:
            f = facturas[0]
            factura = fciva_repo.by_comprobante(f.TIPO, f.LETRA, f.PTOVTA, f.CPBTE)
            if factura:
                print(
                    f"✓ Factura {factura.LETRA}{factura.CPBTE:04d}: "
                    f"${factura.GRINS} (con IVA: ${factura.GRINS + factura.IVAINS})"
                )

            # Detalles de la factura
            fcestad_repo = repos.fcestad1()
            detalles = fcestad_repo.by_comprobante(f.TIPO, f.LETRA, f.PTOVTA, f.CPBTE)
            print(f"  ├─ Renglones: {len(detalles)}")
            for det in detalles[:3]:
                print(
                    f"  │  └─ {det.COD1}: {det.CANT} x ${det.PVTA} = ${det.IMPTE}"
                )

    finally:
        db.close()


def ejemplo_stock():
    """Ejemplo: Consultas de Stock."""
    db: Session = get_session()

    try:
        repos = RepositoryFactory(db)
        stock_repo = repos.stock()

        # Artículos bajo stock
        criticos = stock_repo.criticos()
        print(f"✓ Items con stock crítico: {len(criticos)}")
        for item in criticos[:5]:
            print(f"  ├─ {item.COD1}: Stock actual={item.STUNID}, Mínimo={item.STMIN}")

    finally:
        db.close()


if __name__ == "__main__":
    print("=== EJEMPLOS DE USO - REPOSITORIO CRUD ===\n")

    print("1. CRUD de Clientes:")
    try:
        ejemplo_cliente_crud()
    except Exception as e:
        print(f"⚠ {e}\n")

    print("\n2. Parámetros Globales:")
    try:
        ejemplo_parametros()
    except Exception as e:
        print(f"⚠ {e}\n")

    print("\n3. Cuentas Corrientes:")
    try:
        ejemplo_ctascte()
    except Exception as e:
        print(f"⚠ {e}\n")

    print("\n4. Facturación:")
    try:
        ejemplo_facturacion()
    except Exception as e:
        print(f"⚠ {e}\n")

    print("\n5. Stock:")
    try:
        ejemplo_stock()
    except Exception as e:
        print(f"⚠ {e}\n")

    print("\n=== FIN ===")
