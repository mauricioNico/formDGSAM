package dgsam;

import java.util.Arrays;
import java.util.List;

public interface ListProvider {

    List<String> grados();
    List<String> espBasica();
    List<String> espAvanzada();
    List<String> escalafon();

    List<String> provincias();
    List<String> estadosCiviles();
    List<String> conyugeGrados();

    static ListProvider demo() {
        return new ListProvider() {
            @Override public List<String> grados() {
                return Arrays.asList(
                        "Brigadier General", "Brigadier Mayor", "Brigadier",
                        "Comodoro Mayor", "Comodoro", "Vicecomodoro",
                        "Mayor", "Capitán", "Primer Teniente", "Teniente", "Alférez",
                        "Suboficial Mayor", "Suboficial Principal", "Suboficial Ayudante",
                        "Suboficial Auxiliar", "Cabo Principal", "Cabo Primero", "Cabo",
                        "Voluntario de Primera", "Voluntario de Segunda"
                );
            }

            @Override public List<String> espBasica() {
                return Arrays.asList("Comunicaciones", "Sanidad", "Mantenimiento", "Meteo", "Sistemas", "Infantería");
            }

            @Override public List<String> espAvanzada() {
                return Arrays.asList("Radar", "Redes", "Servidores", "Pronóstico", "Taller", "Instrumental");
            }

            @Override public List<String> escalafon() {
                return Arrays.asList("Técnico", "Profesional", "Especialista");
            }

            @Override public List<String> provincias() {
                return Arrays.asList(
                        "Buenos Aires", "CABA", "Catamarca", "Chaco", "Chubut", "Córdoba", "Corrientes",
                        "Entre Ríos", "Formosa", "Jujuy", "La Pampa", "La Rioja", "Mendoza", "Misiones",
                        "Neuquén", "Río Negro", "Salta", "San Juan", "San Luis", "Santa Cruz", "Santa Fe",
                        "Santiago del Estero", "Tierra del Fuego", "Tucumán"
                );
            }

            @Override public List<String> estadosCiviles() {
                return Arrays.asList("Soltero/a", "Casado/a", "Unión convivencial", "Divorciado/a", "Viudo/a");
            }

            @Override public List<String> conyugeGrados() {
                return grados();
            }
        };
    }
}
