import React, { useState, useRef } from 'react';
import {
  AppBar,
  Toolbar,
  Typography,
  Container,
  Box,
  Button,
  Chip,
  Alert,
  LinearProgress,
  Stepper,
  Step,
  StepLabel,
  Paper,
  Snackbar,
  FormControlLabel,
  Switch,
  Grid,
  Table,
  TableHead,
  TableBody,
  TableRow,
  TableCell,
  TableContainer
} from '@mui/material';
import {
  CloudUpload,
  Compare,
  Download,
  CheckCircle,
  Settings,
  Analytics
} from '@mui/icons-material';
import { createTheme, ThemeProvider } from '@mui/material/styles';

const theme = createTheme({
  palette: {
    primary: { main: '#1976d2' },
    secondary: { main: '#dc004e' }
  },
  typography: { h4: { fontWeight: 600 }, h6: { fontWeight: 500 } }
});

const steps = [
  'Carregar Planilhas',
  'Comparar Dados',
  'Exibir Diferenças'
];

export default function App() {
  const [activeStep, setActiveStep] = useState(0);
  const [filesNucleo, setFilesNucleo] = useState([]);       // [{ file, hasHeader }]
  const [filesUnidades, setFilesUnidades] = useState([]);   // [{ file, hasHeader }]
  const [loading, setLoading] = useState(false);
  const [diffs, setDiffs] = useState([]);
  const [snackbar, setSnackbar] = useState({ open: false, msg: '', type: 'info' });

  const nucleoRef = useRef();
  const unidadesRef = useRef();

  const showMsg = (msg, type = 'info') => setSnackbar({ open: true, msg, type });

  const handleFiles = (e, setter) => {
    const arr = Array.from(e.target.files)
      .filter(f => f.name.endsWith('.xlsx'))
      .map(f => ({ file: f, hasHeader: true }));
    if (!arr.length) return showMsg('Selecione ao menos um .xlsx', 'error');
    setter(arr);
  };

  const handleCompare = async () => {
    if (!filesNucleo.length || !filesUnidades.length) {
      return showMsg('Selecione arquivos em ambos os lados.', 'error');
    }
    setLoading(true);
    setActiveStep(1);

    const formData = new FormData();
    filesNucleo.forEach(item => {
      formData.append('files_nucleo', item.file);
      formData.append('headers_nucleo', item.hasHeader);
    });
    filesUnidades.forEach(item => {
      formData.append('files_unidades', item.file);
      formData.append('headers_unidades', item.hasHeader);
    });

    try {
      const resJson = await fetch('http://localhost:5000/upload-json', { method: 'POST', body: formData, mode: 'cors' });
      if (!resJson.ok) throw new Error((await resJson.json()).error);
      const { diffs: records } = await resJson.json();
      setDiffs(records);
      setActiveStep(2);

      fetch('http://localhost:5000/upload', { method: 'POST', body: formData, mode: 'cors' })
        .then(r => r.blob())
        .then(blob => {
          const url = URL.createObjectURL(blob);
          const a = document.createElement('a');
          a.href = url;
          a.download = 'comparacao_verbas.xlsx';
          a.click();
          URL.revokeObjectURL(url);
        });

      showMsg('Comparação concluída!', 'success');
    } catch (err) {
      showMsg(`Erro: ${err.message}`, 'error');
      setActiveStep(0);
    } finally {
      setLoading(false);
    }
  };

  const resetProcess = () => {
    setActiveStep(0);
    setFilesNucleo([]);
    setFilesUnidades([]);
    setDiffs([]);
    if (nucleoRef.current) nucleoRef.current.value = '';
    if (unidadesRef.current) unidadesRef.current.value = '';
  };

  const tableHeaders = diffs.length
    ? Object.keys(diffs[0]).filter(h => h !== '_merge')
    : [];

  return (
    <ThemeProvider theme={theme}>
      <Box sx={{ minHeight: '100vh', bgcolor: '#f5f5f5' }}>
        <AppBar position="static">
          <Toolbar>
            <Analytics sx={{ mr: 2 }} />
            <Typography variant="h6" sx={{ flexGrow: 1 }}>
              Comparador de Verbas
            </Typography>
            <Button color="inherit" onClick={resetProcess} sx={{ textTransform: 'none' }}>
              Nova Comparação
            </Button>
          </Toolbar>
        </AppBar>

        <Container sx={{ py: 4 }}>
          <Stepper activeStep={activeStep} sx={{ mb: 3 }}>
            {steps.map(label => (
              <Step key={label}><StepLabel>{label}</StepLabel></Step>
            ))}
          </Stepper>

          {activeStep === 0 && (
            <Paper sx={{ p: 3 }}>
              <Typography variant="h5" gutterBottom>
                <CloudUpload sx={{ mr: 1 }} /> Upload de Planilhas
              </Typography>
              <Grid container spacing={2}>
                <Grid item xs={6}>
                  <Typography>Planilha 1 (Núcleo)</Typography>
                  <Button component="label" variant="contained" startIcon={<CloudUpload />}>
                    Selecionar Núcleo
                    <input
                      type="file"
                      hidden
                      ref={nucleoRef}
                      onChange={e => handleFiles(e, setFilesNucleo)}
                      accept=".xlsx"
                      multiple
                    />
                  </Button>
                  <Box sx={{ mt: 1, display: 'flex', flexWrap: 'wrap', gap: 1 }}>
                    {filesNucleo.map((item, i) => (
                      <Box key={i} sx={{ display: 'flex', alignItems: 'center', gap: 1 }}>
                        <Chip icon={<CheckCircle />} label={item.file.name} color="success" />
                        <FormControlLabel
                          control={<Switch checked={item.hasHeader} onChange={e => {
                            const cp = [...filesNucleo];
                            cp[i].hasHeader = e.target.checked;
                            setFilesNucleo(cp);
                          }} />}
                          label="Cabeçalho?"
                        />
                      </Box>
                    ))}
                  </Box>
                </Grid>

                <Grid item xs={6}>
                  <Typography>Planilha 2 (Unidades)</Typography>
                  <Button component="label" variant="contained" startIcon={<CloudUpload />}>
                    Selecionar Unidades
                    <input
                      type="file"
                      hidden
                      ref={unidadesRef}
                      onChange={e => handleFiles(e, setFilesUnidades)}
                      accept=".xlsx"
                      multiple
                    />
                  </Button>
                  <Box sx={{ mt: 1, display: 'flex', flexWrap: 'wrap', gap: 1 }}>
                    {filesUnidades.map((item, i) => (
                      <Box key={i} sx={{ display: 'flex', alignItems: 'center', gap: 1 }}>
                        <Chip icon={<CheckCircle />} label={item.file.name} color="success" />
                        <FormControlLabel
                          control={<Switch checked={item.hasHeader} onChange={e => {
                            const cp = [...filesUnidades];
                            cp[i].hasHeader = e.target.checked;
                            setFilesUnidades(cp);
                          }} />}
                          label="Cabeçalho?"
                        />
                      </Box>
                    ))}
                  </Box>
                </Grid>
              </Grid>
              <Box sx={{ mt: 3 }}>
                <Button variant="contained" startIcon={<Compare />} onClick={handleCompare} disabled={loading}>
                  Comparar Agora
                </Button>
              </Box>
            </Paper>
          )}

          {activeStep === 1 && (
            <Paper sx={{ p: 3, mb: 3 }}>
              <Typography variant="h5" gutterBottom>
                <Compare sx={{ mr: 1 }} /> Processando
              </Typography>
              <LinearProgress />
            </Paper>
          )}

          {activeStep === 2 && (
            <Paper sx={{ p: 3, mb: 3 }}>
              <Typography variant="h5" gutterBottom>
                Diferenças Encontradas
              </Typography>
              <TableContainer>
                <Table size="small">
                  <TableHead>
                    <TableRow>
                      {tableHeaders.map(h => <TableCell key={h}>{h}</TableCell>)}
                    </TableRow>
                  </TableHead>
                  <TableBody>
                    {diffs.map((row, ri) => (
                      <TableRow key={ri}>
                        {tableHeaders.map(col => (
                          <TableCell key={col}>
                            {col === 'Diferença' ? Number(row[col]).toFixed(2) : String(row[col])}
                          </TableCell>
                        ))}
                      </TableRow>
                    ))}
                  </TableBody>
                </Table>
              </TableContainer>
            </Paper>
          )}

          <Snackbar
            open={snackbar.open}
            autoHideDuration={6000}
            onClose={() => setSnackbar(s => ({ ...s, open: false }))}
          >
            <Alert severity={snackbar.type}>{snackbar.msg}</Alert>
          </Snackbar>
        </Container>
      </Box>
    </ThemeProvider>
  );
}
