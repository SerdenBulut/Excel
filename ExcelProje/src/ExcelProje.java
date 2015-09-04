import java.awt.GridLayout;
import java.awt.event.ActionEvent;
import java.awt.event.ActionListener;

import javax.swing.JButton;
import javax.swing.JFrame;
import javax.swing.JLabel;
import javax.swing.JPanel;
import javax.swing.JTextField;

public class ExcelProje extends JFrame {

	private JLabel txtHat;
	private JLabel txtUc;
	private JTextField edtHat;
	private JTextField edtUc;
	private JLabel txt›cerik;
	private JButton btnListele;
	private JPanel pnlUst;
	public static int sayfaBelirleyici = 0;

	public ExcelProje(String title) {
		super(title);
		setDefaultCloseOperation(JFrame.EXIT_ON_CLOSE);

		pnlUst = new JPanel();
		pnlUst.setLayout(new GridLayout(3, 2));
		txtHat = new JLabel("Hat Tipi:");
		edtHat = new JTextField();
		edtHat.setText("AVEA POS");

		txtUc = new JLabel("UÁ Sistem(Sayfa Ad˝):");
		edtUc = new JTextField();
		edtUc.setText("MW TEST03");

		btnListele = new JButton("Listele");

		pnlUst.add(txtHat);
		pnlUst.add(edtHat);
		pnlUst.add(txtUc);
		pnlUst.add(edtUc);
		pnlUst.add(btnListele);

		add(pnlUst);
		txt›cerik = new JLabel("›Áerik burada gˆsterilecek!");
		add(txt›cerik);
	}

	public static void main(String[] args) {

		ExcelProje excelProje = new ExcelProje("AVEA");
		excelProje.setSize(500, 500);
		excelProje.setVisible(true);
		excelProje.setLayout(new GridLayout(2, 1));
		ExcelVeriOkuma excelVeriOkuma = new ExcelVeriOkuma();
		excelProje.btnListele.addActionListener(new ActionListener() {

			@Override
			public void actionPerformed(ActionEvent e) {
				// TODO Auto-generated method stub

				sayfaBelirleyici = excelVeriOkuma.sayfaSecme(excelProje.edtUc
						.getText().toString());
				excelProje.txt›cerik.setText(excelVeriOkuma.veriOkuma(
						sayfaBelirleyici, excelProje.edtHat.getText()
								.toString()));
				excelVeriOkuma.veriDegistir(excelProje.txt›cerik.getText()
						.toString(), sayfaBelirleyici);

			}
		});
	}
}
