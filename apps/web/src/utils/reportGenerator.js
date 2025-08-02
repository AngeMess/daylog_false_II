import jsPDF from 'jspdf';
import autoTable from 'jspdf-autotable';
import * as XLSX from 'xlsx';

// Configuración de colores modernos y estilizados
const COLORS = {
  primary: '#01426A',
  primaryGradient: '#01426A',
  primaryLight: '#2B5F8A',
  secondary: '#FFD700',
  secondaryLight: '#FFED4A',
  accent: '#4A90E2',
  text: '#2C3E50',
  textLight: '#5A6C7D',
  success: '#27AE60',
  warning: '#F39C12',
  error: '#E74C3C',
  lightGray: '#F8F9FA',
  mediumGray: '#E9ECEF',
  darkGray: '#6C757D',
  white: '#FFFFFF',
  border: '#DEE2E6',
  shadow: 'rgba(0,0,0,0.1)'
};

// Función para obtener datos adicionales del usuario según su rol
const getUserAdditionalData = (user) => {
  const baseData = {
    projects: [],
    activities: [],
    performance: {
      completedTasks: 0,
      activeProjects: 0
    },
    weeklyPerformance: {
      workHours: 52,
      extraHours: 12,
      compensatoryHours: 6
    },
    workTeams: []
  };

  // Simular datos según el rol del usuario
  switch (user.daylogRol.toLowerCase()) {
    case 'supervisor':
      return {
        ...baseData,
        projects: [
          { name: 'Creación de sistema web', status: 'Por Hacer', size: 'Mediano', area: 'Tecnología' },
          { name: 'Implementación CRM', status: 'En Progreso', size: 'Grande', area: 'Tecnología' },
          { name: 'Migración de datos', status: 'Completado', size: 'Pequeño', area: 'Tecnología' }
        ],
        performance: {
          completedTasks: 24,
          activeProjects: 5,
          teamSize: 8
        },
        workTeams: [
          { name: 'Marvel end Game', type: 'Feature Team', area: 'Recursos Humanos - Seguridad Informática', supervisor: 'yanose' },
          { name: 'Los aurores', type: 'Agile Product Team', area: 'Tecnología - Soporte Técnico', supervisor: 'pi peta777' }
        ]
      };
    
    case 'portafolio':
      return {
        ...baseData,
        projects: [
          { name: 'Gestión Portfolio A', status: 'En Progreso', size: 'Grande', area: 'Gestión' },
          { name: 'Análisis de recursos', status: 'Por Hacer', size: 'Mediano', area: 'Planificación' },
          { name: 'Optimización procesos', status: 'En Progreso', size: 'Grande', area: 'Mejora continua' }
        ],
        performance: {
          completedTasks: 18,
          activeProjects: 3
        },
        workTeams: [
          { name: 'Fantastics', type: 'Feature Team', area: 'Tecnología - Registro Contable', supervisor: 'Pere wil' },
          { name: 'Los aurores', type: 'Agile Product Team', area: 'Tecnología - Soporte Técnico', supervisor: 'pi peta777' }
        ]
      };
    
    case 'admin':
      return {
        ...baseData,
        projects: [
          { name: 'Administración del sistema', status: 'Activo', size: 'Continuo', area: 'Tecnología' },
          { name: 'Gestión de usuarios', status: 'Activo', size: 'Continuo', area: 'Administración' },
          { name: 'Backup y seguridad', status: 'Activo', size: 'Crítico', area: 'Seguridad' }
        ],
        performance: {
          completedTasks: 42,
          activeProjects: 6,
          systemUptime: '99.8%'
        },
        workTeams: [
          { name: 'Marvel end Game', type: 'Feature Team', area: 'Recursos Humanos - Seguridad Informática', supervisor: 'yanose' },
          { name: 'SysAdmin Core', type: 'Technical Team', area: 'Tecnología - Infraestructura', supervisor: 'admin001' }
        ]
      };
    
    default: // Empleado
      return {
        ...baseData,
        projects: [
          { name: 'Desarrollo módulo ventas', status: 'En Progreso', size: 'Mediano', area: user.mainAreaArea },
          { name: 'Testing funcionalidades', status: 'Por Hacer', size: 'Pequeño', area: user.mainAreaArea },
          { name: 'Documentación técnica', status: 'En Progreso', size: 'Pequeño', area: user.mainAreaArea }
        ],
        performance: {
          completedTasks: 15,
          activeProjects: 2
        },
        workTeams: [
          { name: 'Fantastics', type: 'Feature Team', area: 'Tecnología - Registro Contable', supervisor: 'Pere wil' }
        ]
      };
  }
};

// Generar reporte individual en PDF
export const generateUserReportPDF = (user) => {
  const doc = new jsPDF();
  const userData = getUserAdditionalData(user);
  
  // Header con logo real de Banco Cuscatlán
  // Fondo blanco para el header
  doc.setFillColor(255, 255, 255);
  doc.rect(0, 0, 210, 35, 'F');
  
  // Agregar logo real de Banco Cuscatlán
  try {
    // La imagen se carga desde el directorio público
    const logoUrl = '/Banco Cuscatlan logo.png';
    // Mantener proporciones originales del logo (aproximadamente 4:1 ratio)
    doc.addImage(logoUrl, 'PNG', 20, 10, 60, 15); // x, y, width, height - Proporciones correctas
  } catch (error) {
    console.warn('No se pudo cargar el logo, usando texto como fallback');
    // Fallback: texto si no se puede cargar la imagen
    doc.setTextColor(COLORS.primary);
    doc.setFontSize(16);
    doc.setFont('helvetica', 'bold');
    doc.text('BANCO CUSCATLÁN', 20, 20);
  }
  
  // Línea separadora sutil
  doc.setLineWidth(0.5);
  doc.setDrawColor(200, 200, 200);
  doc.line(20, 32, 190, 32);
  
  // Título del reporte
  doc.setTextColor(COLORS.text);
  doc.setFontSize(18);
  doc.setFont('helvetica', 'bold');
  doc.text(`REPORTE DE ${user.daylogRol.toUpperCase()}`, 20, 40);
  
  // Fecha
  const currentDate = new Date().toLocaleDateString('es-ES', {
    day: '2-digit',
    month: '2-digit',
    year: 'numeric'
  });
  doc.setFontSize(12);
  doc.setFont('helvetica', 'normal');
  doc.text(`Fecha: ${currentDate}`, 140, 40);
  
  // Información personal
  let yPos = 60;
  
  doc.setTextColor(COLORS.text);
  doc.setFontSize(12);
  doc.setFont('helvetica', 'bold');
  
  const personalInfo = [
    ['Nombre:', user.fullName],
    ['Puesto:', user.daylogRol],
    ['País:', user.country],
    ['Área:', user.mainAreaArea],
    ['CuscaID:', user.cuscaId],
    ['Jefe Inmediato:', user.inmediateBoss || 'No asignado'],
    ['Estado:', user.isActive ? 'Activo' : 'Inactivo']
  ];
  
  personalInfo.forEach(([label, value], index) => {
    doc.setFont('helvetica', 'bold');
    doc.text(label, 20, yPos + (index * 8));
    doc.setFont('helvetica', 'normal');
    doc.text(value, 70, yPos + (index * 8));
  });
  
  yPos += personalInfo.length * 8 + 15;
  
  // Sección de proyectos asignados
  if (userData.projects.length > 0) {
    doc.setFont('helvetica', 'bold');
    doc.setFontSize(14);
    doc.text('Proyectos Asignados:', 20, yPos);
    yPos += 10;
    
    // Tabla de proyectos
    const projectColumns = ['Proyecto', 'Estado', 'Tamaño', 'Área'];
    const projectRows = userData.projects.map(project => [
      project.name,
      project.status,
      project.size,
      project.area
    ]);
    
    const tableConfig = {
      startY: yPos,
      head: [projectColumns],
      body: projectRows,
      theme: 'striped',
      styles: {
        lineColor: [200, 200, 200],
        lineWidth: 0.1,
        cellPadding: 4,
        fontSize: 9,
        textColor: [1, 66, 106], // Azul corporativo del banco
        halign: 'left'
      },
      headStyles: {
        fillColor: [1, 66, 106], // Azul corporativo del banco #01426A
        textColor: [255, 255, 255], // Texto blanco
        fontSize: 9,
        fontStyle: 'bold',
        halign: 'center'
      },
      bodyStyles: {
        fillColor: [255, 255, 255], // Fondo blanco
        textColor: [1, 66, 106], // Texto azul corporativo
        fontSize: 9
      },
      alternateRowStyles: {
        fillColor: [245, 245, 245] // Gris muy claro para alternancia
      },
      margin: { left: 20, right: 20 }
    };
    
    autoTable(doc, tableConfig);
    
    // Obtener la posición Y final de forma segura
    yPos = (doc.lastAutoTable && doc.lastAutoTable.finalY) ? doc.lastAutoTable.finalY + 15 : yPos + 60;
  }
  
  // Sección de Rendimiento Semanal
  if (yPos > 230) {
    doc.addPage();
    yPos = 30;
  }
  
  doc.setFont('helvetica', 'bold');
  doc.setFontSize(14);
  doc.text('Rendimiento:', 20, yPos);
  yPos += 5;
  
  doc.setFont('helvetica', 'normal');
  doc.setFontSize(12);
  doc.text('Esta semana', 20, yPos);
  yPos += 10;
  
  // Tabla de rendimiento semanal
  const weeklyPerformanceData = [
    ['Horas laborales', `${userData.weeklyPerformance.workHours}h`],
    ['Horas extra', `${userData.weeklyPerformance.extraHours}h`],
    ['Horas compensatorias', `${userData.weeklyPerformance.compensatoryHours}h`]
  ];
  
  autoTable(doc, {
    startY: yPos,
    head: [['Tipo', 'Horas']],
    body: weeklyPerformanceData,
    theme: 'striped',
    styles: {
      lineColor: [200, 200, 200],
      lineWidth: 0.1,
      cellPadding: 4,
      fontSize: 9,
      textColor: [1, 66, 106], // Azul corporativo del banco
      halign: 'left'
    },
    headStyles: {
      fillColor: [1, 66, 106], // Azul corporativo del banco #01426A
      textColor: [255, 255, 255], // Texto blanco
      fontSize: 9,
      fontStyle: 'bold',
      halign: 'center'
    },
    bodyStyles: {
      fillColor: [255, 255, 255], // Fondo blanco
      textColor: [1, 66, 106], // Texto azul corporativo
      fontSize: 9
    },
    alternateRowStyles: {
      fillColor: [245, 245, 245] // Gris muy claro para alternancia
    },
    columnStyles: {
      0: { cellWidth: 80, halign: 'left' },
      1: { cellWidth: 40, halign: 'right' }
    },
    margin: { left: 20, right: 20 }
  });
  
  yPos = (doc.lastAutoTable && doc.lastAutoTable.finalY) ? doc.lastAutoTable.finalY + 15 : yPos + 80;
  
  // Sección de Equipos de Trabajo
  if (userData.workTeams.length > 0) {
    if (yPos > 230) {
      doc.addPage();
      yPos = 30;
    }
    
    doc.setFont('helvetica', 'bold');
    doc.setFontSize(14);
    doc.text('Equipos de trabajo:', 20, yPos);
    yPos += 10;
    
    // Tabla de equipos de trabajo
    const teamColumns = ['Equipo', 'Tipo', 'Área', 'Supervisor'];
    const teamRows = userData.workTeams.map(team => [
      team.name,
      team.type,
      team.area,
      team.supervisor
    ]);
    
    autoTable(doc, {
      startY: yPos,
      head: [teamColumns],
      body: teamRows,
      theme: 'striped',
      styles: {
        lineColor: [200, 200, 200],
        lineWidth: 0.1,
        cellPadding: 4,
        fontSize: 9,
        textColor: [1, 66, 106], // Azul corporativo del banco
        halign: 'left'
      },
      headStyles: {
        fillColor: [1, 66, 106], // Azul corporativo del banco #01426A
        textColor: [255, 255, 255], // Texto blanco
        fontSize: 9,
        fontStyle: 'bold',
        halign: 'center'
      },
      bodyStyles: {
        fillColor: [255, 255, 255], // Fondo blanco
        textColor: [1, 66, 106], // Texto azul corporativo
        fontSize: 9
      },
      alternateRowStyles: {
        fillColor: [245, 245, 245] // Gris muy claro para alternancia
      },
      margin: { left: 20, right: 20 }
    });
    
    yPos = (doc.lastAutoTable && doc.lastAutoTable.finalY) ? doc.lastAutoTable.finalY + 15 : yPos + 60;
  }
  
  // Métricas de rendimiento
  if (yPos > 250) {
    doc.addPage();
    yPos = 30;
  }
  
  doc.setFont('helvetica', 'bold');
  doc.setFontSize(14);
  doc.text('Métricas de Rendimiento:', 20, yPos);
  yPos += 10;
  
  const performanceData = [
    ['Tareas Completadas', userData.performance.completedTasks.toString()],
    ['Proyectos Activos', userData.performance.activeProjects.toString()]
  ];
  
  // Agregar métricas específicas por rol
  if (userData.performance.teamSize) {
    performanceData.push(['Tamaño del Equipo', userData.performance.teamSize.toString()]);
  }
  if (userData.performance.portfolioValue) {
    performanceData.push(['Valor del Portfolio', userData.performance.portfolioValue]);
  }
  if (userData.performance.systemUptime) {
    performanceData.push(['Uptime del Sistema', userData.performance.systemUptime]);
  }
  
  autoTable(doc, {
    startY: yPos,
    head: [['Métrica', 'Valor']],
    body: performanceData,
    theme: 'striped',
    styles: {
      lineColor: [200, 200, 200],
      lineWidth: 0.1,
      cellPadding: 4,
      fontSize: 9,
      textColor: [1, 66, 106], // Azul corporativo del banco
      halign: 'left'
    },
    headStyles: {
      fillColor: [1, 66, 106], // Azul corporativo del banco #01426A
      textColor: [255, 255, 255], // Texto blanco
      fontSize: 9,
      fontStyle: 'bold',
      halign: 'center'
    },
    bodyStyles: {
      fillColor: [255, 255, 255], // Fondo blanco
      textColor: [1, 66, 106], // Texto azul corporativo
      fontSize: 9
    },
    alternateRowStyles: {
      fillColor: [245, 245, 245] // Gris muy claro para alternancia
    },
    columnStyles: {
      0: { cellWidth: 80, halign: 'left' },
      1: { cellWidth: 40, halign: 'right' }
    },
    margin: { left: 20, right: 20 }
  });
  
  // Footer
  const pageCount = doc.internal.getNumberOfPages();
  for (let i = 1; i <= pageCount; i++) {
    doc.setPage(i);
    
    // Agregar logo de fondo como marca de agua en la parte inferior derecha
    try {
      const backgroundLogoUrl = '/Logo de fondo .png';
      // Posicionar parcialmente fuera del área visible para mostrar solo una parte
      doc.saveGraphicsState();
      doc.setGState(new doc.GState({opacity: 0.25})); // Opacidad aumentada para ser más visible
      doc.addImage(backgroundLogoUrl, 'PNG', 145, 220, 120, 120); // x, y, ancho, alto - Logo más hacia adentro para que salga más
      doc.restoreGraphicsState();
    } catch (error) {
      console.warn('No se pudo cargar el logo de fondo para la marca de agua:', error);
    }
    
    // Texto del footer
    doc.setFontSize(8);
    doc.setTextColor(COLORS.text);
    doc.text('DayLog - Sistema de Gestión de Proyectos', 20, 285);
    doc.text(`Página ${i} de ${pageCount}`, 170, 285);
  }
  
  // Guardar el PDF
  const fileName = `reporte_${user.daylogRol.toLowerCase()}_${user.fullName.replace(/\s+/g, '_')}_${currentDate.replace(/\//g, '-')}.pdf`;
  doc.save(fileName);
};

// Generar reporte individual en Excel
export const generateUserReportExcel = (user) => {
  const userData = getUserAdditionalData(user);
  const currentDate = new Date().toLocaleDateString('es-ES');
  
  // Crear libro de trabajo
  const workbook = XLSX.utils.book_new();
  
  // Hoja 1: Información del usuario
  const userInfoData = [
    ['REPORTE DE ' + user.daylogRol.toUpperCase()],
    ['Fecha:', currentDate],
    [''],
    ['INFORMACIÓN PERSONAL'],
    ['Nombre:', user.fullName],
    ['Puesto:', user.daylogRol],
    ['País:', user.country],
    ['Área:', user.mainAreaArea],
    ['CuscaID:', user.cuscaId],
    ['Jefe Inmediato:', user.inmediateBoss || 'No asignado'],
    ['Estado:', user.isActive ? 'Activo' : 'Inactivo'],
    ['Email:', user.email],
    [''],
    ['MÉTRICAS DE RENDIMIENTO'],
    ['Feedback Score:', userData.performance.feedback],
    ['Tareas Completadas:', userData.performance.completedTasks],
    ['Proyectos Activos:', userData.performance.activeProjects]
  ];
  
  // Agregar métricas específicas por rol
  if (userData.performance.teamSize) {
    userInfoData.push(['Tamaño del Equipo:', userData.performance.teamSize]);
  }
  if (userData.performance.portfolioValue) {
    userInfoData.push(['Valor del Portfolio:', userData.performance.portfolioValue]);
  }
  if (userData.performance.systemUptime) {
    userInfoData.push(['Uptime del Sistema:', userData.performance.systemUptime]);
  }
  
  const userInfoSheet = XLSX.utils.aoa_to_sheet(userInfoData);
  
  // Aplicar estilos básicos
  userInfoSheet['!cols'] = [{ width: 25 }, { width: 30 }];
  
  XLSX.utils.book_append_sheet(workbook, userInfoSheet, 'Información');
  
  // Hoja 2: Proyectos asignados
  if (userData.projects.length > 0) {
    const projectsData = [
      ['PROYECTOS ASIGNADOS'],
      [''],
      ['Proyecto', 'Estado', 'Tamaño', 'Área'],
      ...userData.projects.map(project => [
        project.name,
        project.status,
        project.size,
        project.area
      ])
    ];
    
    const projectsSheet = XLSX.utils.aoa_to_sheet(projectsData);
    projectsSheet['!cols'] = [{ width: 25 }, { width: 15 }, { width: 15 }, { width: 20 }];
    
    XLSX.utils.book_append_sheet(workbook, projectsSheet, 'Proyectos');
  }
  
  // Guardar el archivo Excel
  const fileName = `reporte_${user.daylogRol.toLowerCase()}_${user.fullName.replace(/\s+/g, '_')}_${currentDate.replace(/\//g, '-')}.xlsx`;
  XLSX.writeFile(workbook, fileName);
};

// Función para generar ambos reportes
export const generateUserReports = (user, format = 'both') => {
  try {
    if (format === 'pdf' || format === 'both') {
      generateUserReportPDF(user);
    }
    
    if (format === 'excel' || format === 'both') {
      generateUserReportExcel(user);
    }
    
    return {
      success: true,
      message: `Reporte${format === 'both' ? 's' : ''} generado${format === 'both' ? 's' : ''} exitosamente`
    };
  } catch (error) {
    console.error('Error al generar reporte:', error);
    return {
      success: false,
      message: 'Error al generar el reporte: ' + error.message
    };
  }
};
