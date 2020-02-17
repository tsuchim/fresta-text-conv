# fresta-text-conv
Converter for Fresta-Texts

- Clone Program Tree

 $ git clone git@github.com:tsuchim/fresta-text-conv.git
 or
 $ git clone https://github.com/tsuchim/fresta-text-conv.git

- Update submodules
 $ git submodule update

* If you have "Internal Server Error" caused by permission denied on linux with SELinux,
  to set context manually would help you:
  $ chcon -t httpd_sys_script_exec_t /var/www/fresta-text-conv/data/text/update.cgi
  
