# Prompt Maestro: Migración Sistema Stella (VBA a Arquitectura Fullstack)

Actúa como un Arquitecto de Software Senior y Desarrollador Fullstack. Tu objetivo es migrar un sistema de gestión de producción textil (Stella v3.8) basado en Excel/VBA hacia una aplicación web moderna.

## 1. Stack Tecnológico

- **Backend:** Java 21, Spring Boot 3.3+, Spring Data JPA, Hibernate, QueryDSL, Spring Security (JWT).
- **Frontend:** React 18 (Vite), Tailwind CSS, Lucide React, Axios, TanStack Query.
- **Base de Datos:** PostgreSQL.
- **Infraestructura:** Docker & Docker Compose (Multi-staging).

## 2. Lógica de Negocio Crítica (Extraída de VBA)

### A. Gestión de Calendario y Capacidad (Modulo 5)

- Implementar un servicio de **Días No Laborales** que calcule festivos (incluyendo la Ley Emiliani de Colombia) y fines de semana.
- La capacidad estándar de una operaria es de **528 minutos/día**.

### B. Distribución Automática de Carga (Modulo 3)

- Al crear una asignación, si la operaria no tiene minutos suficientes hoy:
  1. Calcular el remanente del día actual.
  2. El excedente debe distribuirse automáticamente en los siguientes días hábiles disponibles.
  3. Crear registros de `AsignacionDiaria` vinculados para cada día hasta agotar la cantidad.

### C. Control de Stock por Tarea (Modulo 3)

- Para cualquier `Lote`, `Tarea` y `Talla`, la suma de `cantidad_asignada` en todas las asignaciones no puede superar la `cantidad_recibida` del lote.

### D. Cálculo de Avance Ponderado (Modulo 7)

- Un lote tiene un avance basado en **Unidades Equivalentes**.
- `Unidades_Equivalentes = (Suma de unidades completadas en TODAS las tareas del lote) / (Total de tareas de la referencia)`.
- El estado del lote cambia a `COMPLETADO` solo cuando el avance es >= 100%.

### E. Proyección de Entrega (Modulo 7)

- `Fecha_Fin_Estimada = f_inicio + (Unidades_Pendientes * Tiempo_Total_Referencia / Capacidad_Planta_Dia)`.
- El cálculo debe saltar automáticamente días festivos y domingos.

## 3. Modelo de Datos (PostgreSQL)

Diseña el esquema ER con:

- `operarias`: id (uuid), nombre, minutos_diarios, estado (enum), especialidad.
- `referencias`: id (string pk), nombre, descripcion, tiempo_total_seg.
- `tareas`: ref_id (fk), num_tarea (int), nombre, tiempo_seg, dificultad.
- `lotes`: id (string), ref_id (fk), color, talla, cantidad_recibida, avance_equiv, estado.
- `asignaciones_diarias`: fecha (date), operaria_id (fk), lote_id (fk), tarea_id, cantidad, unidades_completas, estado.
- `festivos`: fecha (date pk), descripcion.

## 4. Requisitos de Infraestructura (Docker & Ambientes)

Crea los siguientes archivos de configuración:

1. **`Dockerfile.backend`**: Multi-stage (build con Maven, run con JRE 21).
2. **`Dockerfile.frontend`**: Multi-stage (build con Node, serve con Nginx).
3. **`docker-compose.yml`**: Definición de servicios (app, db, frontend).

### Manejo de Ambientes

- **DEV:** Perfil de Spring `dev`, base de datos con `ddl-auto: update`, logs en `DEBUG`.
- **QA:** Perfil `qa`, base de datos persistente, validación de datos.
- **PRD:** Perfil `prod`, seguridad JWT estricta, optimización de assets en React, base de datos con backups volumétricos.

## 5. Instrucciones para la IA (Pasos a seguir)

1. **Persistencia:** Genera el SQL inicial y las Entidades JPA con validaciones de Hibernate.
2. **Services:** Implementa la lógica de `PlanningService` para la distribución de carga.
3. **Security:** Configura Spring Security con JWT y roles para `SUPERVISOR` y `ADMIN`.
4. **Frontend:** Crea el Dashboard usando `Tailwind CSS` y una tabla de `Trazabilidad de Lotes` que muestre el progreso visual de cada tarea.
5. **DevOps:** Genera los archivos Docker y un script `.env` de ejemplo para cada ambiente.
