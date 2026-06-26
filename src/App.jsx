import { useState, useEffect, useRef } from "react";
import * as XLSX from "xlsx";
import { initializeApp } from "firebase/app";
import { getDatabase, ref, onValue, push, remove, update } from "firebase/database";

const FIREBASE_CONFIG = {
  apiKey: "AIzaSyAIJBHt3j73Ustm1BIxA8329yzFS2j1uMM",
  authDomain: "petit-subcon.firebaseapp.com",
  databaseURL: "https://petit-subcon-default-rtdb.firebaseio.com",
  projectId: "petit-subcon",
  storageBucket: "petit-subcon.firebasestorage.app",
  messagingSenderId: "1089448725745",
  appId: "1:1089448725745:web:63769e226cb991a3f4ad50"
};

const USERS = {
  "3588": { name: "정태식", role: "master" },
  "5447": { name: "이원호", role: "master" },
  "0441": { name: "조홍휘", role: "master" },
  "8238": { name: "김세윤", role: "staff" },
  "1998": { name: "박형수", role: "staff" },
  "3828": { name: "장한이", role: "staff" },
  "3452": { name: "원민섭", role: "staff" },
};
const LOGO = "data:image/png;base64,iVBORw0KGgoAAAANSUhEUgAABDUAAAKLCAYAAADmYLPeAABhYklEQVR42u39+ZNc1bk3en5LpVJpRAiEhBAIsMxgsA3G83DOef3evt0RHdH33+24cbvb5/jYxhjPeMBgARaDQAgJgdCApv7hWfvNlJDQUJXDyvp8InZkSYjKrL1z79rrm8961tL27dsDAAAA0JtNdgEAAADQI6EGAAAA0CWhBgAAANAloQYAAADQJaEGAAAA0CWhBgAAANAloQYAAADQJaEGAAAA0CWhBgAAANAloQYAAADQJaEGAAAA0CWhBgAAANAloQYAAADQJaEGAAAA0CWhBgAAANAloQYAAADQJaEGAAAA0CWhBgAAANAloQYAAADQJaEGAAAA0CWhBgAAANAloQYAAADQJaEGAAAA0CWhBgAAANAloQYAAADQJaEGAAAA0CWhBgAAANAloQYAAADQJaEGAAAA0CWhBgAAANAloQYAAADQJaEGAAAA0CWhBgAAANAloQYAAADQJaEGAAAA0CWhBgAAANAloQYAAADQJaEGAAAA0CWhBgAAANAloQYAAADQJaEGAAAA0CWhBgAAANAloQYAAADQJaEGAAAA0CWhBgAAANAloQYAAADQJaEGAAAA0CWhBgAAANAloQYAAADQJaEGAAAA0CWhBgAAANAloQYAAADQJaEGAAAA0CWhBgAAANAloQYAAADQJaEGAAAA0CWhBgAAANAloQYAAADQJaEGAAAA0CWhBgAAANAloQYAAADQJaEGAAAA0CWhBgAAANAloQYAAADQJaEGAAAA0CWhBgAAANAloQYAAADQJaEGAAAA0CWhBgAAANAloQYAAADQJaEGAAAA0CWhBgAAANAloQYAAADQJaEGAAAA0CWhBgAAANAloQYAAADQJaEGAAAA0CWhBgAAANAloQYAAADQJaEGAAAA0CWhBgAAANAloQYAAADQJaEGAAAA0CWhBgAAANAloQYAAADQJaEGAAAA0CWhBgAAANAloQYAAADQJaEGAAAA0CWhBgAAANAloQYAAADQJaEGAAAA0CWhBgAAANAloQYAAADQJaEGAAAA0CWhBgAAANAloQYAAADQJaEGAAAA0CWhBgAAANAloQYAAADQJaEGAAAA0CWhBgAAANAloQYAAADQJaEGAAAA0CWhBgAAANAloQYAAADQJaEGAAAA0CWhBgAAANAloQYAAADQJaEGAAAA0CWhBgAAANAloQYAAADQJaEGAAAA0CWhBgAAANAloQYAAADQJaEGAAAA0CWhBgAAANAloQYAAADQJaEGAAAA0CWhBgAAANAloQYAAADQJaEGAAAA0CWhBgAAANAloQYAAADQJaEGAAAA0CWhBgAAANAloQYAAADQJaEGAAAA0CWhBgAAANAloQYAAADQJaEGAAAA0CWhBgAAANAloQYAAADQJaEGAAAA0CWhBgAAANAloQYAAADQJaEGAAAA0CWhBgAAANAloQYAAADQJaEGAAAA0CWhBgAAANAloQYAAADQJaEGAAAA0CWhBgAAANAloQYAAADQJaEGAAAA0CWhBgAAANAloQYAAADQJaEGAAAA0CWhBgAAANAloQYAAADQJaEGAAAA0CWhBgAAANAloQYAAADQJaEGAAAA0CWhBgAAANAloQYAAADQpc12ASzcOb2cZCkVWg6P17uS5HLbhq8BAAC6GwAB82trku1tW02ykmRLO3dX27YlFWRsusnj9a4kuZrkUq4NNy7eYDvbHi8kOd/+fN5hAQAA5oFQA+bD1iQ7k+zKKMTYOfb19vZvVsa21ev+/nZDjfHqjOHxRoHGhYxCjPFQ42yS00nOJPm0PZ6Nag8AAGDKhBowXctJdqfCi3vGvt459jiEGDva47aMwovldt4OlRrDf1/JjQONpRu8hqsZBRqX2teft68vta8vZxRiXGjb50nOtb//JMlnqUDjTPv6bCrkOJtR4DGEHgAAAOtOqAGTMQQW97avh8qLXe3vx//bzrH/viMVVgzTTra1Pw/BxdKUf46rGVVufN62C6lwY6jkuH4bgo1TqYqO00k+SgUhZ9rjJ+3vr3qrAAAAd0uoAetjKcn+JHuT7GuPe9vf7cm11Ri7x7Yh0Ng8xz/XlrbtuMW/vZprp6gMocbJJCfan4cw4+MkHyZ5tz1+lApMAAAAbn/Asn37dnsB7t6OJA8meTQVYOxLcmDs6weT3J/RdJONGCReTYUaJ1OhxqlUqPF+krfb48kkx9u/O54KRQAAAL6UUAPu8JxJ8kAqtDjQvt6b5FBGlRn7Mqra2GKX3dDHqfDivYwqNT7MKNw4muSD9vVpuwsAALjhAE2oAbd0b0aBxcFcG1zsSXJf+/N97c/b7LLbdjWjaSrD44kkx1KhxvD1qYwCj5OpJqcAAMAGJ9SAG7s3o2kkB9vXB1IVGUMVxt5UXwzVGOvrXKpyY+jFcaw9fpCq7Hi/Pb6dqvgAAAA2KKEGXOuBVHDxeHs8mFGQcTDJwxFiTNv5VIXGsSTvtMe32jb8+VhqhRYAAGADEWpANfs8kAoshhBjPNR4OFW5wex9nAowjiZ5M7V6yjBVZTz4sFQsAABsAEINNrK9SQ4n+UqSh9o2hBtDoGHZ4/l0MRVefJBR742h8ejRVNjxXirkAAAAFpRQg41oWIL1qSRfT1VlDA1A96aWXqUvH6eCjWGZ2LeT/CvJ66mKjiNRvQEAAAtHqMFGsicVYDyV5Mkkz2QUalixZHF8kqrUeDPJq6lg49W2vW/3AADA4hBqsBHsSYUYj6WmmzyT5Gvt73baPQvrXKpC40hGocbr7c/CDQAAWABCDRbZE6kqjMcyCjUeT/J0Em/8jePzVNXGMBXlrVSwMaya8rZdBAAAfRJqsIh2JnkhyfOpAONw2w4m2Wr3bFhXU81Fh6kpR9v2epLXUj04AACAjiyvrKzYCyyS7Ul+nOTfkvx7kh8m+U6S+2Ilk41uKcnu1Oo2X0nyQNt2psKupdQKKgAAQCeEGiySlST/kQoz/i0VaOy3W7iBLUn2pap3dqTCsF2pgONiqtkoAAAw54QaLIrdSX6S5KepSo3vRRNQbm1LkodSwcY97X20N9Vc9nKS03YRAADML6EGi+Bgkv+RCjT+R6qfhiVauV1LqSDjYHt8sG33p8KNq0lO2U0AADCHN/MahdK5nUn+jyT/99R0kycX9Of8PDUt4lJ7vJqqJLgy9m82JVlug/TNY3+30v5+pf03bu29jJaBfTXJK0n+muRDuwYAAOaHxon07ulUmPHTJIc6e+1Xk5xt2/kkF1KBxbCdv+7P49uV3DzUuP5xJaNwY9i2JllN9ZIY31wTykNtO9S2fUkOpAKOv7djAwAAzJgBDL07lOQbmd9A42ySz9rjufb1Z0nOtO1Gocbl3DjEGP/z9YFGUlUYS/lisLG5/f0QaGxO9ZLY2rZtqYqX7e3vt7d/N4QeW9t/35FqprmR5qx9NRVm7E2tmvJokkeS/CO1FOwVpyAAAMyOUIPe7c58rHDySdvOtNDi0yQfJ/ko1Wzy7Nh2pv334d+eS00vudDCiktjocWltl0d+3r4b1czmk6yNDbA3tT+PB5sDH83XrWxpQUX21pgsTr259WMmmfekwozdrdtT2qJ3Pvan3cs+HtsR5IfJHksyRNJDqfCjj8leTlWSgEAAKEGdOhckuNJPkhyLNVM8vTYdrL99+tDjSHQuDTH14XVVJCxJ6MwYwgx9qQqF/a2v9vXvt6fxQ44hgaij6SmptzX9tNv23EGAABmMHiBnp1p4cG0nExVX5xIhRnHU80j32v/7eOMQo1T7d9d7GyfDhUhnyV5v/3dpiT35tpQY6jW2J/kgVS48XB7HAKQRXQotTLKvampOfcm+X2q3wYAADBFQg16906SI0meT/V/WC+fp0KKoeLiVCqwONG+HgKLk237MKNA4/IC7ucrYz/rm+3v7mkD+vvb9kCqgmF/KvR4IKOKjv2pyo9FsSPJj9o+OJBaDnZ/kj+29wAAADAFQg16dyRV/r8/yfeztuqAs6mg4lgqLDne/jyEF6cyCjrOtMdPs3GbRQ59RI62P+/OaCrKfRkFGuMVHPtTUzgeWJB98PX28xxsP9+D7f14xKkJAACTJ9Sgd8eT/CLV+PLTJC8kefwO/v8rqakjx5O8m5pS8k4bqB/LKNg4mQoyuLlh2s3r7c/DVJX9GfWhONi+fiQVdDy2AD/33iT/kapW2Z8KdH6R5BVvCQAAmKzllZUVe4HeHcuoGeeF1NSRpdTqHtcHd5dT4cfx1Kfpf0j1Q/hD2/7YHn+X5M+pwON0+57cmQupKTnvpoKOISQaepJ8mAqLLmS0xGzP9qVCm2FJ3M3t57/grQAAAJOhUoNF8beMpoS82waXu9vgcljW9EqS86nwY2jkOTT7HBp/Dj0zWF8Xk7w1th1MVW48nFFPikfb1/tSvSp6tDfJT1Orouxuf/5TTEcBAICJWNq+fbu9wCLZmeRwahrAzoxCjaSqNC5ktKzqEGwM1QJM165UiPFgO16H2p8PpKYQHUqFHT26kqr6+V2S36SCjT97nwEAwPoSarDoVlJVGklVC2y0pp67Uk05dybZluRqqqLlo1RlyrzYngoz9qcqOA637clUM877O93/7yX5VSrYeDkVcnzqtAQAgPWhpwaL7kqqQuNyG9BvFNuTfCPJt9rj11pA8JVU9cPeVA+LjzIfS9BeTFXNvJ3qvTG+jO6ZVE+T7emv78au1HSazamKoaX2s513agIAwNoJNWDx7EzykyT/nuTHqRVhnk0FG4dTK4/sT/V82JTqcTFPzufapqKnUr1Ojrf/ttJ+xl4MVSjbUg1RVzIKawAAgDUQasDi+UGS/2fb/iPJM6mlUx9J9an4aqqPxb1tgH0utYztvDmbmr7xr/b4bkZhx+n22nd3ckxW2zG4P6NeLyfbzwEAANwloQYslgNJ/rck/68k32mD5+stpaZFHExNyfk0yRstRJhHF1JBxmupZWCHFWrOJLmU5L6b/JzzaG+qSmY1Ne3n47YBAAB3QagBi+XrSf73JP/H7Zz/qWqNk6kqiKMd/HynUgHMx6lQ41zbdqaCmh7sSgUbV1PBxqmo2AAAgLsi1IDFsSU19eT/kVoS9XZsS1VovN3Cgh6WHL2aqtZ4LxVufJrks1Qz0Z3po5no9lTVxrDc8MdRsQEAAHdMqAGL48Ek303y0zZgvl1XUn0r3k1N8+jF+VR1ycepqo3T7eurqV4bm+f89e9KsicVbFxtP88Jb2MAALh9Qg1YHI+lQo2f5M5WB1lp4cCx1EooVzr7uYfpMx/n2sqNrZn/RqL3pAKoTe3Pl5J84K0MAAC3R6gBi+NrSb6d5Ee5syqFLakGnO+nKjU+7vBnv5DR6iinU8HGuVQVxL5Uc9R5tSsVbAyv8Ww7HgAAwC0INWAxbE8t3fp8atWTO/VJCwTeSfWq6NVnqf4gn2XUZ+N8qnJl+xy/7l2pqo1LqTDmk+ixAQAAtyTUgMWwJ8mzSb6ZCjfu1MVUpcabSY50vi+upipOPkpVPZxJBRtbU8u/zqvdqaVez6XCmI9SwQwAAHATQg1YDA9mFGocvov/fyU15eGNJK8syD75NNUA9ZO2nW9//1DmdzrKsCLKlVQ4MwQzAADADWy2C2AhbB/b7vb/352aBrGc5PKC7JfPk/whtarIJxk1Ef1RqpfIPPp2e91XU9NRfpnklLc4AAB8kVADFsP21PSKtZRe7WzbahavOuDtVKAxLP16NrVSzANz+FqXkvwwtSLK5VTA8X+lQg4AAGCM6SfQv02pKSdPplZAefAuv8+ZJP9I8qcs5pSH86npNZ+kVkv5LBXs7pvD17qU5OH2+i61Y/OetzoAAFxLpQb0b2uqumIlNXXkbg3TV7Yt+P76cyq0OZWq2riQmvIxj76TCmE+SK1Mc8zbHQAARoQa0L+VsW3TGr7PWvty9OSfqeqHsxn1D5nHYGNzKtj4MFWp8X9G41AAALjmhhno23iosXUN32dnkh2pqo+N4P0k/5VaaeRKKtz43hy+zvuS/DijlVB+luqzAQAAG55QA/q3KaNAYy1NclYyaji6UXyc5KWMKlw2pSoj5s1jqWDj0yTnkvzc2x4AAIQasAiGQGM1a6+y2J61BSM9Op7kV6mGnGmPP5jD1/nt1JSZT9prftVbHwCAjU6oAf3b2rb1qLLYiKFGUj0r/jMVaFxsfzePwcYPUg1O3001Dz3l7Q8AwEYm1ID+DZUaQ7XGWmzUUCOpJV5fSq0gM0zF+eacvcbVJM+nlt59I8mL3v4AAGxkQg3o30pGU0/WWqkxLA27UX2c5Dft2riSqtx4Yc5e42NJvp5aweV4kiNOAQAANiqhBvRvCDXWo8HnRlnS9cucSK2K8nlqtZGLSb4/Z6/xG0mOphqHXmxfAwDAhiPUgP4Ny7mux1KsW7NxlnT9MqeT/CK11OuWJHuSPDlHr+9QkudSocbZ9qi/BgAAG45QA/q30gbe6zFtZL0qPhbBmVSPjV1J9ie5L8neOXp9z7XXeCrJsSS/dcgAANhohBrQv9VUqLEeYcTWCDXGnU712Hggyc4k/54KN+bBjlTT0BOp6SevpkIOAADYMDbZBdC9YerJyjp+r2W79X95OxVs/CpVuXF+jl7bviRPJDmc5HGHCgCAjUalBvRvZWxbj+819NU4a9f+L7/PqInqniQ/mKPX9nSqSuOJJH9N9QEBAIANQagB/RsPItZqqPhYsVu/4KVUqHFf2+alceiuJM+mlnY9kuRPDhUAABuFUAP6tinrW6mRCDVu5mKSPyY5mORAaurHvXPy2p7IKNg4nmocCgAAC0+oAX1bbtt6BhGbI9S4meOpSoj9Se5P8tMkS3PwunYleSbVMFSoAQDAhiHUgL4tpQKI9QwihqCEG/tzqq/GjtR0nR/Pyet6NrUSysmMVkMBAICFJtSAvi3n2iko68H0ky93KckrqSVedyTZneTrc/LavpGq0vhbhBoAAGwAQg3o26aMpp+sV3WFUOPWjqf6a+xK9dZ4JvOxRPZ9qYqNJ1L9Pj52qAAAWPQBEdCvYdrJek4XEWrcnn+llnr9Y2op1XnxTCrYOOwQAQCw6IQa0Lf1bhKa9r1W7drb8npqmsff5uw98ViSrzg8AAAsOqEG9H8OTyLU2GrX3paLqWDjldRyqvPiUKpS4zGHCACARR8QAf0app+o1Jidf6SmoPwqySdz8poeT/J0ahqKqUQAACwsoQb0fw5vyvqHGgbCd+aVJC8leXGOXtOTqf4aTzg8AAAs8oAI6Nfy2LZehBp37mgq1Phlkt/NyWt6MrXU7DNReQMAwIKypCv0bdPYtl6GnhorqZ4R3J7fJrknyYOp6oh7Zvx67k/yfGqVlleT/MUhAgBgEQdEQL+WMplKjU3r/D03imGJ11fn5PU80bZDDg0AAItIqAHc6Lqw7PpwVz5OBRrzshLKttQKKA8l2enwAACwiIMXwLk8bgg0VGrcndeT/DPJO3Pyeg6mQo29Dg0AAAZCwDxZyvoHEMvRLHQt3k/1sfjXnLye/Un2RagBAMACEmpA/+fweocaQ6PQ7XbvXXsrVa1xYg5ey5ZUqPGAwwIAwCIOiIB+LadWMVrPc3lLkh0RaqzFP5O80rZ58GCSA0n2ODQAACwSoQb0fw4PwcZ6WUmyGqHGWryd5LVU09A35+D17ElVauxzaAAAWLQBEdCvYTnX9TyXN0VPjbW6kpqCcqQ9ztruVE8NoQYAAAtFqAF9Wxrb1stKhBrr4c0kbyQ5mtn31tiTCjUOJLnfoQEAYFEINaBvQ6Cx3ku6DtUfS3bxXTuTmobyTpIPZvxadqZCjf2p/hoAALAQhBqwGCZVqbHZrl2To6lQ49gcvJYh1NifagYLAADdE2pA/+fwevfU2JJqFLpq8Ltmx1M9Nd5NcnHGr2VfkoOpKSh7HRoAABZlQAT0awg0ltfxe64k2do2fTXW7niqp8ZHM34de8e23Q4LAACLQKgB/Z/D612pMSzpuhqhxno42bZTM34d96Qaht6XZJfDAgDAogyIgH5NYqUSlRrr61TbPpqD13J/qkpju8MCAMAiEGpA34YqjfVs6LkSy7qup7NJTqeqNWZtd2olFKEGAAALQagB/Z/DwxKs60mosX4upkKNE0k+nfFr2Z2aeiLUAABgYQZEQL+GSo2ldf6+4301WLuTST7M7Jd23Z5RtQYAAHRPqAH9n8PLE/i+KjXW11CpcXwOXssw/WSbwwIAwCIMiIB+LU/oPBZqrK8zqWahHyT5bMavZZh+stVhAQCgd0IN6P8cnkSlxmqsfrKePk2tfnKibbO0PfpqAACwQAMioF9LUanRg7NJPklVa8x6FZTtGYVWAADQNaEG9G0pemr04EJqCsrpts3SEGhoAgsAQPeEGtC3pQl932HQK9hYP2dT01DOzPh1DCvbqNQAAKB7Qg3gZgPflUyuEelGdH5sm4djK6wCAKB7BivAlw18lzOZ6S0b0YVUoHExyaU5OLZCDQAAuifUAOfwrQa+rhPr42KSz9vj5Rm+js0RagAAYEAEzIFJTQ8ZDzVUaqyPi0mutG3WhBoAACwEoQb0bVLTQ1ZimsJ6u5zZVmjc6PgCAEDXhBrQt0kNTleSbMloBRTW7tLYdXfW114NYAEAWAhuaqFvkxqcbopKjfV2dezrpTl4PUsOCQAAvRNqQN8mNf1kOdVQ0if6i3vtveowAADgxhqYpaUJfl+Bxvoa9udyZl8lMetlZQEAYF0YsED/5/Akl3WdVCXIRrR5bF/OOtSYp6alAACwpgER0K9Jhg5DPw3XicXbn1dS1RoAANA1gxXo11LbJnUer0aj0PW0te3TrXPwWi5GqAEAwAIQakC/ljOdSg3TT9bH9ratzvh1XIpQAwCABSHUgL7P30n31DD9ZP3satv2Gb+Oi0nOJ7ngkAAAsAiDIqBfk2w4OQQaKjXWbmeSPW3bPePXMgQa5x0WAAB6J9SAfk16+snwvV0n1m5XKtC4r22zdLZtQg0AALpnsAJ9n7+Tnn6iUmN9DJUae9s2S0OoYfoJAAALMSgC+rScyU4/2RQ9NdbLrlSYcSDJjhm/ljOpUOOcwwIAQO8MVqDv83eSlRTDyicqNdZuZ2rayQNz8FqGqSefOywAACzCoAjo06R7XqxET431MvTU2DsHr+VcauqJJV0BAOiewQr0ff5O8hxeTrI5KjXWw+4k9ye5Z8av42KSz1KVGkINAAAWYlAE9Gmoopjk9JPNrhNrtjuj5Vxn7ZOMemoAAED3DFag7/N30tNPhiko3L1hxZP75+C1nErycSrYAACAhRgUAX2aRr+LIdjg7h1Msj/Jvjl4LR8mOZHkpMMCAMAiEGpAv5baNslKikn37Vh0K6llXPelVj+ZpU9SgcaJJKcdGgAAFoHBCvBlLOm6NvvbNg+rnpxOVWicjFADAIAFIdSAvi1N+DxesovX5KGMKjVmbajS+DDVWwMAALon1IB+LWfyocM0+nYsqp0Z9dO4bw5ez6kkx1PBxiWHBwCARWCwAv1azuSnh5h+cvcea9vBzH7lk3NJPkhyrG0AALAQhBrQr81tm6TlKTzHIlpJ8mSSp1LBxqz34Ykk7yZ5K8n7Dg8AAItCqAF9D5wnPT1kc1Rq3I1HkjzRtsfm4PWcyqhSAwAAFoZQA/o1BBqTXtJVX407dyDJV5J8NcmWOXg9J1PVGscdGgAAFomBCvBlhkDDteLOHEpyuD3Og2Hlk08cGgAAFomBCnCra4RKjTuzOdUc9NE5eT2fZxRqAADAwt18A9zMNKa4LJpDSR7PfPTSSKqPxrCUKyySnanlkncnWU1yJbXSz+n2nrd0MQBsAEIN6Nc0poVsTjUkda24fV9LNQidl332Zts0CWUR7E5VQu1P8kBqueSd7Tp1Jcn5jEKNoTnue0mu2nUAsJgMVKBfm5IsTeE5TD+5fV9J8o1UP4158FFqGdd3k3zm8NCxvamw8PFUNdS+VKBxb5Kt7X7mcpILST5NNcc9ngo1jiY5kuSfduNE7WjHZKieWUlysR2PU+2YnLabAFhvQg3o1/KUnsP0k9uzP8n3kzzfBl7z4PU2mHvH4aFjTyR5IcmzGQUbe9vAeXsbPC+lQo2LqSkon7RB9AepSqVX2zn6m/ZvWD8H2zF5JMmD7djsab83LiT5OKO+Pu+3Y/JeKugAgDUTakDfJl1FsRLTT27HjiQ/SvLvqWBjaQ5e0/tJ/tq2txwiOvVcO6d+1L5+MhVk3OyeZjU1HeWBjCqm/tW+3t/+35fbQJu12Z2qTHs6yVOpSrWDGQVOy6npQKcyWlL6/VTl2Jup0PXVJGftyrk03FtcsSuAeWegAv2aRhXF8tjGzT3dBl4/yfw0CH09yd/aoOGcQ0SHnkmFGT9p290ukfxoarrKMCViNRVsfGAX37WDqeqZFzIKNr7a9u31DrTHz3JtqPFwKgD5Q5IP7dKZ29u2nRlVQCVVbXM2oyldQihg7gg1oF/T6HcxfH89NW5uS7uh/26Sr8/JazqX5LVUlcarDhGdDpq/m+SHbTu0xu+3LcmP2zVtpV3Tfh4VG3d7bP6tHZfvpKYF7b6N/29HqmLmcKqqY38bRN+TmhZ01K6duvtSQfzDGfWp2dXOl82pBrvnU0HGqVSPmreTvBFBFDBHhBrQr6UphA2mn9zak6lPKr89R6/pSKpK45VY9YH+7ExVAPwwVQG1nj1qftCum1dSDSx/ZnffkR2pcOh/a49fu8vvMz6IHgbSL7XrFpO3LRUafrP9DnskFXDsaeff1vbvrmZUqTGEGkdTYflQCXjS7gRmzUAF+jWN6SfD9zf95OaDr6dz+59UTstbqUqN9xwiOvSNVKjxfBtwrbfvphqJvpPRyijcnm+mgqGftGvfWmxpx2J3G0QPq6W8bjdP1O5UYPjTJN9rx/HB2/x/P2/nzOFUELIvyW+jGTUwY0IN6JfpJ7P3bBuAPTNHr+kTAzU6dqidU89P8Lxaaufua6lPm50rt+dgOy7fzdoDjXFPpgKNK6klX4+mqgNYf/elgoz/SIUa373D/39LqnfKgdTUoV3tPuHzVL8NgJkQakDfljLZKorNMf3kZr6RKo3/Tmp++Lw40gZrbzlEdOjpdm59ow2YJuVAannYw0n+lFqdg5tbTQUaz7VtvT2emuJwLKOpc6yvB1IVGj9OrdT13TV8rx2pap3VVBh1Psl/pprBAkydT1+hX8uZ/NKh06gG6dHBdmP/nVSZ/Lz4JFW6/VaseEJ/draQ4emMlmOd9Hn8cHvkyz2dqpx5NpMLm55OBU2P2d0TObe+nQo0fpK1BRrjnm/bC0m+ZTcDs2KgAs7fLzONvh09ejajMuwH5+h1/bVtyunp0f7U9JPDU3q+B1IVGw/Z9V9qXypwGEKHSf6+eTyjpV5ZP8+kQo0fpqafrNfv9JX2vYcpY8/a1cAsKCmHvk06bFgZ2yjPpT6VWu955Wv1l9SyiL9NBRvQ4+D5UNZ3tZMvc19GwcZ9sYrDzTyeUaixf8LPdSCjfg2mBK2Ph1MNXl9IVRdunsAx+0ZqNaGPU1NQ3rLbgWkSakC/plE9YfrJtQ61G8Pn2k3cvHg/NQf990n+GMu40qfdLVyY5jX0ntQylvdHqHEj29ug9fFMJ2zaPbaxPg4neSpVUbFtQs/xVGr644dtO5rqtQEwFQYqwK1u+k0/KQ+kPuX6bqrMds8cvbZ/Jfl7KtCwjCs9D6DvmfJzrqaaHhpE3/yY7E31HTkwpeOxvW2s3b2pHiWHM9lpXUtJvpaafvJkqjoEYGqEGtD3+Tvp5VY3R6VGUsvYfT/VYO1Hma9pJ6eTvJEKNf7stKBjW2cwmB2ec1cm33i5RyupJpPTqqBZbs+5atevi4dSFYaPZfLTSHelgo2vRvNdYAaDIoCbWcrkg5MefCfJD1JN1p6bs9f2r4yWcYXe70mm3b9nJaNgY5tDcNNjMs2wSWXg+nkg1atmWo1XH04FKRq9AlP/ZQX0ee5O88ZvI3+C+dVUhcaPU8HGPDmZ5G+pKg2hBr1bnsGAdpjqsDWqA+blmCT6MayHHalw4b5Mbhne662mgpT77X5g2gMjoO8bzmncWG7UUOOxJP97kv9bkv+Ys9d2LslLSX6d5A9JzjsdWIB7kmlXamxvg79tscrTzZxPcnaKz3dxys+3qIZw4Z5Mt9Lm/lSQssMhAKZ5AwH0Z9p9LjZiOfBXkvyPtv0o8xfs/Dm1fOsfUpUa4Lp254Ylq015uLGzqb49HyS5PKXn/Cy1PChrc0+qH8pqphvYbW/Pq9krMDVCDej33N00pefZiI1CDyf5adt+nPn7xOlEagnXPyX5i9OBBbmmrWT61RKXpzhY79HJJMeSvNW2aTzfyVSQwtpsSwUaq5luKL91bAOY2k0EwJddI2ZREj5LX0vyP1PTTn6a+ezi/kqqUuPPbQAAvRt6Wkz7WnOlbYKNm++fI0n+muTVJJ9P+PleT/J2KrhlbYaQcMuUn3e1PafpXMDUbLYLgNu4Mdoo14qvp1Y4+Y8k/5bq4j5vPkiFGn9pAwBYBENV2LSngajUuLXXkxxo20NJvjWh5zmRCk/eTHLGbl+zpcx2SpcxBjA1LjjArcyq+/20DYHGj9rjPAYal1N9NH6fCjZgUVxt27RXvViJvhq340+p5o8PJnm0fb2eziZ5OVV9dsTu7tr4OQUwFUIN6NdSprO066YFv1ZsTvKdJN9N8v0k30s1CZ1Hv0zy30lejPJsFsvV9jjtqoldbRuWduXGTqdWW9qTCjb+t6xfz4RPUqs4vZgKbd+yu9fFcqa//HsymvJiiWRgqjfzALe6MVrUTzC3pCozvp/khVS4Ma+Bxu9TgcYvk7zmbcmCuZoKNC5N+XlXM2pq6JPlL3csFTocSAVA30py7xq/5wdJ/tiua79qX7M+ZhFojN8zGGMAU+OCA9yOpbZdXaCf6d5UdcZPUmHG85nPKSdJ8l7qU8yft0dYxGtMZnCNGf80W6hxa68k2ZcKgc4keSrJI6mVNu7UP1O9gV5u17WXU8u5sj7n0/De3jTD5weYCqEG9Gl5LGiYtE1ZvFDjkVR1xvdT/TOez/wt2zr4PMlvkvxX2zQ1ZBHtaAPjLTN47q2pyoPdDsNteTnV++REkndSS2B/JclXb/P/fz8VaPw91avj96mpLVfs2nW9R5hlbwuhBjBVQg3gdgyfZC7CTedTqeqMYcrJC5lOOHS3/tBu+F9McsFbkQW1t4UKs+hrcU+q+uBgezzucHypT1JTRU6kpqQcTS3D+l6Sh9ux3HXddfVS+/f/SvJGRqHG31LVHwKNyZnFikKWSQamSqgB/ZrWQHyWJazr6d5URca3k/w4yXOZ3/4Zg1fa4OFXqU9EYRGtJDmUZH8LGKbtYKrK4GRq+sPv2mB9EtfrRal2u9CuT++mlmAdgoqHktyfqrzZmqq8uZrkXJIP2789kgo3/pmq2mAywcIsw4XLEcIDUyTUgD5tmnLQsLltvd6kfC21qsm3U300nsvdzQGfpldS001+ngo1YBHtTvJ0kmeTPD6j13BvuzasppYqfTQVapxJhbpbMqpUG6bjbbruepyMPhEfrs1Xc234PCxZez7Jp0lOtVDgeKbfIHW9nGzbP1JVGvsyWk1mtW1L7Wc+ngpnj7afn8mZ5dSPzRlNkQWY2oUH4MsMN+g9Xi+GDv3fS1VnfKcNWObdPzJqDPqLjgc8sJLRCiPbxwY829rgd0+qJ8NzSZ6c4eu8L8m/tYH5U22gfqFd+1baa778JQPITdc93mhAN3xqfja1ROqJNsA/mlrG9Egb/Pfo09R0kr+PHfeVsX13vv3cTO+8G9//s/jdu9VhAKZFqAHcyrA8W28rA+zLqDLjh6k+Gjs7eN3vpXpo/KJtJ7wFmfPAYhjAjAcYK6kKh9Wxf7O9/d2W9vXOVKXGwVS1xqxXH1pKTUn7SipIvNL+bnNGjZKHaovh6/Fr5O26cF2o8WaS15P8NdVjYhGmml1sG7OxO9dWzEzbnvY7GGAqhBrQr/Gb6knrqVJjZ+oT32dTlRnfTjUF7eH1fzQWZvwi6z+vH+4muNjZBke72tdbx0KK8VBj9bpwY/NYqLGaa/ss7ExVa+xINZZ8oIP7o2EVqLVO+1ttA759SZ5JVWkcTq3KdCDJH1OrgsDdOtjeS/sym1V9Hm6/h7+Z5M8OBzCLX9oA48bLqufZ9tSc/GdSn/p+LTX15OlO9vNnqQqNX7XHN731mIHVFjIcyGhFkmFVkiHYGMKLmz2Ol71vzijk2Nb++xBsUB4b29/3tf29I7V0qmoH7tQjqZDsUHtf3T+D13AoNaXs0/a+PpGqULqYmop05Qb3GTcyfu8xVEkNFVTDBzvD9zwbq+iAUAPgFjcV87rm/EqSJ1KfCj2V5Ovt8YlUA8BevJzk1y3Q8MkW07Q9o09197evD2X0Ke94qLF9LLRYzbXTTJbtyru2mlpe+p6MKl82JfllFmfFFCZrRzt3n0yF+YdTFROzei0vtGvCodR0qwupaV2XcvNQY/kGf788FmpczqgB7xByXEh9KPBJRg14P4ilmUGoAXQRNAxhw6Q7jI83ypu3MONw276eqswYKjTu6ex4/jrJf6emnLzs7c0UfaWdM4+1AdCBVOn6sMTqbrtoqr6a0TSdtAGbawJL7Xfe9ZVR28f+fneqKuKx1PLls+5T81Sq8uiZFmpcSVVVjDfcHQKK5evCjVx3j7PpBqFGMlo69vTY9kFq6ua7qea7R719QKgBzP+NzqQt584b4U06zDiUUXXG0D/jqTYg681vk/xnCzR+m5uvsADr7YkkP0qViR8eCzIO2DUzdSDVD+h8qmz/zWgYvKi2pKZlbcsXp2+t5tpVTL4s1FjJaDWhh1ugMQ8rfT2QL++Xc3Wd7mVuFGocbde0V1NNeM94u8HiEmpAv6ZRpTE8zzyEGjvbTdpjbXu6BRpDtUZvLqcqNH6e5P+X6qVhycM7s33sxv/6aVLDJ35Xxvb35VTp8zAH+8IG3nePpRro/ntqhaDDbX8yP8HG821g9opQo3urGfVLuaf9PtuRUcPcrWOP46HG5utCjdUbhBvDdW+YHrYv/VQrrtc9zDBNLkm+0YKNtzPqK7I7ye9SzbgBoQYwZzcD02jeuXlsm4WDqQagh9pA7PG2PZlqiNajy0n+K1Wh8bNUlQY3t5z6BHJYpnBn23a1gcGW64KNpetCjaGZ3KUkn7cw42xq/vWwtOap9ueN0GjuQAsyfpLkp+18Yv4M1WiHUiuiaBran4OpKSDDtK49bbt37Pq1owUaWzJqojsEGMNy6kNwsere/bbsb9tjqcqV3W3fvRS9NkCoAczdwHgaUxWWZvTz7W039UPDs8fbDcqhzK752Xq4lOTFFmb8LNUIkBvbk9F0iGEljj1tMDB8MjcsMTp+4z9esXE5o7nc4xUaQ6AxNJU71f58sn19bEEHkaupTzJfaMGGQGO+HRq75lkRqR972++tJ8aO36H293tSVRvDtYvJHoeftOte2jX9FzEVBYQawNwEGlfyxaZZkzAsmzatT7Dvz+gTyuFxCDMe7Py4fZr6pOjnSf6v9jUjSy3EeKAd67354vKi97VQ457Up507xkKNpVu8j28n1Pi4ff1+knfa40epCo9F8GQLNZ5L8k1vubl3oF3/Dqemoui5M//31V/LaAWup1PB4UOpykJLGU/fSpIftmv4UKHnwwQQagBzYOgCPo0b3En37lhpQcZw0/douxkcppwcymKswPB+aoWTX7bH33sb/6/j/0iqNPuRsQDjgYymnDzQAozda3gvLLUBxTCoGG9edzKj5QA/adtQrTGEGu+nwo/32tc92tPOra+lmuu6B5h/ezOaencwVnKYZwdTYeGwrPiwtOpeu2bmlpJ8L7X06+l2TX/VbgGhBjBbQ6AxVGtM8lxemkCwsZL6RP5g6hP5h1qYMYQYj7fB7dKCHK+jqUagP0tNPXnFWzjb23EeGr0+Nnbch/LsoRHopN3XtnGfpXptDNs7qSUC30ryWpLX09/c7P2pMvjhXGP+bWrvzWHaglBjPj2S5N+SfDtVAfVE5mP1Ea79nfNMu5a/2bYLdgssBqEG9GmaJcibMuq2vlY726DqYHt8vN0MDtvB1NSCRfJmKsj4RapK468b/L27ktGqNTfqmTIvhgZ+w8Dk/Tag/Fd7nY+mwo2/pwKQHuxpg+N9LqFdGa7B7tnm03KqR80PUysKfSN6ZcyrIUgfVkV5yy6BxeAXJPQ7MBxf8m2Stqf6F+xtz3XpLgZS480eD7avH03ylfa4qIOsV1NhxotJXs7GDjSWUyHGMP1h6JfyVHt/zLsH2/bV9p79akbz5v+R5EiqrHmebW3nswFXX86mGhta8nk+PZ1afve7bWO+HWzb/gg1YGEINaBPw+Bk+xSea1eq/HloWPfPL/m3W1K9D4Zy6WHb1x7va18PZfCL/Inx71ONQH+RCjRe36Dv1eX2vhlv/PpMRp+W9ea+1NzsIdg43H6m11Oh1etzPPi8nGqWquS6H1dybTNb5s9jY9c15t8wlWu3XQGLQ6gBfdqdqp6YxlSNPRkFGifb372X5FwbsO5KhRP3t20IMva3vx8CjvtTgceejJbbXESfpgKN36b6aLyU6sWwEQ1BxuHrtqcX4Gfb37avtJ/x9VRI82iSv+XLw79ZOZPRsrX0cz05kerfctzumDurGTVz3WV3dGEINHampnZdsUugf0IN6NMwheP+KTzXttQ81OFTwvtSncMvpabA3JNaSeKBVIgxfD0sybmRPg05kuSPSX7XtpezMT9d3ZdaAeD51BSNoRHoofaeWSRDiDdMSxl6g/w1Nf1ongKtE6kVXY6mmuU97FI6946nqjSOdTb4Gq7/21KVQafbz/D5gh2frW1wbIWTvsY+uzJaClzlGizIiQ305em2DStFTMPhVNn6zjYw/bTdYG9tNwf3p4KMYYC30T6xupiqzvh9KtT4fWqFkwsb9P35XCrQ+FYq1HhsA/zc9yX5SWoln4fb48FU1cZfkpyfg9d4LNW49kg7j4Ua8+9YKhh7p5PXe3+qQusrqUqm1XYdHFYQeivJGwt0fDZlFGzQj+V27IAFIdSAfm6cHmgDkW+kOq0/PcUbqaVUc8d9qcqDc+3vhpu53alP5DbqoONPSX6dCjP+lI257OKOVHXGC217rm1bNth++EoqbNw/9nhPkj+3gd0sXUwFGgcymlJm2cn5dTYVQh3N/DehXUrybKqvxNdTlUv7U5+En09VnBwde//9Yw7Oh/UwNOze5u3alSttu2hXwGIQasB83ywNK4YMA6RHWrjwjcymKdnQN4O6IfpzKsR4ObXCyR+zMefnPpaqzHguybfb+/OxDfze2Nb2w0Nt29POm9+3Qd0svZuqHhlWIjrUBqTMn1dSU5hem/PXuTPJd9r2fDv/D6eCzsFH7b13pP23R1J9h/7Z+TFabttVb9eunG2bfhqwIIQaMJ8OpCoxhrXUh+1QuyE0f3e2PkkFGi+3G/OX5mCwOivfaIOZFzIKNjTMG53He1Ihx642+NuWmo4yS6+217Wvbd9zqObOqxlNY5v3yq8XUlOvftK+fuAG/2YIxB9LhWn3p6q4LqXvZTUvtsHxp34vd2NolnzaroDFIdSA+fJAqlfGE6mKjMdSn/QOZex77KKZe6MNNP6YCjVezsZdleDbSX6U5AcZTYniWluT/DgVauxpj7vb+2ZWTRMvpiqMdrTXdy41dcCgbD68mqr8+nUqPJ1nT6bCzB8m+bfcepnxe1Ih2s5UoPFJajWtXhuInk41cn23/e5m/h1LTX2yChQsEKEGzIcH2oDwqdQc98Pt60dy40+9mL63U2Xgf0mtbPHn1CepG3VO7veS/I/Up7M/NCC+pW9mtPTx0Ez356lPeWc1GPvvVKDxbntPH06FqMNKRhrpTf+Y/L1dV36VCjbenuPXO/Raei4Vam6/g//3mfbe+yjVN+RPnR6zi6lKk1fbz3Sft/Fcu5xRn5oTdgcsDqEGzN7hNih8LvWp1yOpaSZ6V8yHz1KVGX9tA46/tBvYf23gffJckv9IhRo/ysZatnctHmz7bVdGpfe/yuyCjTNJ/rMNyr6S+qT5oVRV2L52XIdtWP5wtT1uz+ItzzvLgfG/2nXlL6lQ46XM/7STYTrkk+29faeeTlW+PZF+Q42kph6+0vbFv6d6bDCf/tbOs7dSFRvAghBqwGztS/LTtn0r1THeQGF+/CsVZvwpozDj76lPGDeqIdD4aapKQ/+MO7MtNR1lOVVyfznJb1Lh2ay81bYHcm2lxn2pqpJh2sxqKozZkgo1drTHLRlVdWxKfYI/LHW5fezfcK1zLbj4V9v//2iDrj+kpjTMu0OpysJDd/n/72hBwOHUFMteB5nHWqhxsL3XvyPYmEtHM+pT81o0CQWhBrAudrfB4f9s2367ZG6cyWiayV9SU01eTc393si+MxZo/LQNWrk7P8ho6tJKqn/CJzN+TR+2LS3IGEKN+1uosXUs2Fhtfx4qODa3gdxS+/OmjCo6to79+5Wxv9+e6q2wdw7vRy6kgqZh4LPpJl9fbo83GiBdHHu83P7N5dQSp2favn67DbaOplYC+WdmV7lzJw6kKnsOtcH8WoKRx9r36vmT81cyqlg7m6pCOeAyNzc+SFXFvdSCjbfsEhBqAGu3muT7GTVZFGjMh3OpUuLXM1pOcdjObfB9842xQOPfBRrr4vupKShXUlUbL7bB9Dw41bYjLXzYldGUk5ttw/KW46HGeJCxOvb1sBrMfW1g+3Tmo3/QB+38fyfV4+JK+1nGP3kf7zVy5brAItf9t1z336+MhRofpYLSd9qA/lRH792DSR5uA/e1VCXsTVUHHUqFx2c6PZdPpJr/Xmw/w7FUBcqhVNWTyo3ZOdbCjBfbMXrFLgGhBrA+XmgDmu9Fx/R5cD6jZm+vpcrA/5qa7/2h3ZNnUlNN/qMFGqacrI8t7Rrweaoq4GxqKsq8OZtbVw8sZVStsZJRxcbKWCiwOvbfd6eqQfa2QON8KjCb5X3Jv9r+f7kFGydbCDH8HFczmlqzfF1gMV6NcX3wcX34MVSBnE6fKyftboP1Ycnxtdo/FpK82vH5/G4LNE6mKm8ea/to6FGzt+27XTHNdFqOpFYqe6md13+KaScg1ADWxdNJvttCjRfsjpm62AKMf7ZBzN9aqPG6MON/eS7VA+J/phqDCjTW145Utda5Nti9lCqP7s3Vdj5dbAHF7djWBsWn2yDv8VRfoVk4kepl8bO2veqteVMHMpo2sh6hxr5UoHGoXYsvdbxvTqcqAt7KKKjZl1GwcV8q2NiZ0bSslVw7pWtLKjTbnNHULu7MsVSw9ErbXk7y22zc1cpAqAGsq02pktRvtE0J/2xcSlVhHGmDl9dbmPFq6tM2yhOp8O3HqalSPQQaF9vA4mwbXF/ItZ+gD+fhUDmwLaOpELP6nbg7ybdTn95fTFVu/GUDvL/OtfNwJaNGkbMKNd5p14A/CjRuaW8brD+U9ak4uK8N+IdqhvcXZFB9rL2X9t4i1Lj+cZjKtaldk1Yz6k0zhB6bM5riNUzz2p6qJNpILrdr/Kftmn8mNYXsWEbVl39NfWAh0AChBrBO9qc+jXoid98xnrv3+diNzpG2/TO1bv2bbnqucSDJ86mqou9kfpvenW7biVTZ96kkH7eb2wv54rSAwfWhxj3tcU9GJeJ72n+fhkfavr6UUc+FtzbIe+1oKkw8McPXcGpsIMSX29MG6A+u0/cbn4p074KEGoMzY+fy0ENmaJK7kmsb7q62bXMLJ4ZqjWG1oW3XBR7JtY13h304VL3MS8BxtZ3bw7X6XPtdPP4zpF2rh2lawzSvpbGf93K7Po5P+fq8XeeHIPvTdi5/0K4rR9LnFC9AqAFzbXdGpbZMd8BypAUXRzLql/FOu/HhWptSlUTfaduTc/b6PmyD4KPthvXEWKgxBBxnM5rOcXXs/x2aOQ49IMYrNXa3QdWeth3IaL7/4Sn8zhz28+ft5ziajTH/+1w7dp/O6PkvZ/Qp7wmn/y3d186N9Wzsurd933sXeL8NAcfNrrkr14UW4393fQXHMPgfwtldbR8eateqZ1JTXR+c4c97NVX59FYqMDw9do2+MBa6DD14rmTUOHm4Rg+hxnCefn7dv7maCoE/aefwZ+3rk5mfpsuAUAMWznDzsceumIpPUlUZQ8+MN9rjq+lrpYFpe7KFGs+3x3kKM4Zw6o12XN9rwcZws3x6jb8T723b8InnwVSvh6fbYOFwJvsJ6JOpSpN3k/y9bRvB0I9jli7niyuYcK2hyeveCQQle1IVUxvR0EB2LQPxzWOBxul2Pm3N7IKiPyb5eepDhHfbNfpE+9179rr31HhQMx7kDqHG0LPn81wbUgMINWBGN4TD/FgmOzg5kmoQ9rexUOOt1CdGfLmvtTDjm5mPpQjfa8fzjba9lVEflI/W8XkuZVT18c9UE8+hlPtIatrYky3geHSCP+/TqeDmqQ0UaszS8Km4ZTdvbWcLH9Z7GfJh6sRuu3hN169/pCqetmZUbTaLUOPzdo3+Y6pB55upKoobOefQAWsl1IDpunLdI+vnXGoe7TtJ3s6oCeg/U0s1vm8X3ZbnUqvyvJDZli6nHcdXU8HUkdR0jLfbNo3VaT5rg4TX2vMfTgUbw9ePZTLTUu5JBRpfTwVzRzbA+255hqHC+Lz9ZdfnL7U7o6aX6+m+9n3vTwUnZ+zqu/ZeCxGeztoq19biRPtd/Ea7fn7usACTJNSA6RrmbZ+Oao31vHl6c2zA+1Z7PNq+Nkf+9j2Y5FupKo2nZvg6ziX5U+pTvr+lypePpqpsPpvB67nabs6Pjr3XDmc0f/3pCeyvr6QqZr6ZCuQ+W/D33lCCPivLGfUv0DD45nalgo1JBFC7U4He9gg11mponDn0FZr2/f6n7TV8HIEGMAVCDZj+L/oTbXB2wO64axdSn0a9l1FVxltj27tR0no3g8qhQuPZVMf9Wfgwo5LlP2a05O48HM9LqcqfU+09djCjZqUXWgCxXu5px+GNtk9+seDvv2EKyCyff3nGr6EH21OVFJOwM6OlS1mbYdWnyzN8/gt+DwPTItSA6RqqCl5Nla7fZ5fckSHIePe67a3UFJONslrEJHwryQ9Sy4p+dUav4cMkv0/ycqpS46/t2M7bjfFHqU9BT6Y+Ub7SBmN7kzy0js/zzdSUqlOp4OS1BX7/XZnhuXs1oyaF3F7wMAnbJ/z9N6rlGZ3Pl6LqCZgSoQZM14WMGlgeSPJTu+RLDY0bj6eqW4YpCMdSYcYHqdL8dw1I1uSeVC+Nb7fHWQ0sX8uoMuPNdtzn9ZO+c+31bc1oqebH2nm9XqujLKWmAr2XUW8PJvf+4/aCh60T/N7DxtpszWg61dIMnn9Te16VT8BUCDVg+l5vA589bSD0rF1yjU/bYPFEKrQ4ngoujqWmmhxvf/9BlLaul8MZLeG6bUav4WRG4dWJjJYknGcXUxUbn6aqKYaVUx5Yx+c4kOrZ8Xg7Nov6nl+a0eArM3zeHu2aYOiwKyo11stKZtt8d9OMz2lggxFqwPSdTZXYD+XqB7Kxp6FcSpXzfzj2OFRhDJUYQ7jxTnyiOomb32dSUx0emuHrGC//X2kDnGFax/mx/35l7PfXytjNc9p74+rYv7l6g/fLjZbuHP7N5fb10g1e2xBibMroE9Bh2cTtqaWah3+z3oaVVvalplktaqgx60ahlnT9cttTVUm7J3gtmuT332hm2admuA6bDgoINWCBHW/BxhBq/DQbqzna+RZgfJDRp9vD10N4cXwszDAvd3K+nqoWenrGr2NnOx8OtuO9MzVd62JuPLVo03XbeLCxdIOAY3zweqMb/etDk/HwY/j7i9f9t02p8GVfaiWU/RMakD3QrhN7FzzUmOWnusrkb20IHLZP+Dl22tULw4cQgFADFtxbSV5qA5Xl1KoTexfw57yUWtbtdNtOjD0O2zD1YHwKgiBjsramliH9fmrayf4Zv55tqYqEtNdy+ksGnVeuGwBfX32xaSycuHTdwHk5X6wKGAKQizcINYbvdXnsPTn+XKupYONAkieS7JjQ/jmwoNcH+jEEDpMONXa18+qCXX7XhuvXrKuPTD8BhBqwAfypDS7PpqZZfKMN7O7p9Of5pG1nxkKMj1NVGScz6jtwun09hB2ftn/D5C23wffjqWkn32vBxjx4ZCzQuJhrp5dsui5gGMKIXPffvyzU2JQbN7BbyqgaY3hcar8jx7/X5etCjeHPWycYZgz2ZfbB06KfF5tiCsqtAofdGU21moShr8YuoUbXrr8mAwg1YMG91Ab1H6QqFJ5ug86vJtkyx6/744yWtDyf5LMWTAxBxRBinGxBxvD1yRZ8MH172/vr6SRPpqaefCvz1dNlS9a30eYiDigN9piFoUpj64SfY5LNSDeK5TkIFYQawNQINWA+/C2jpphHU1NT3ko1B3w0k/8U+E79ObU07VB1MYQaH6bCiyHAGB7POMQztTfV8+GxjPpnDKHGVrunC+PLXQo1mMX7bxphg2Vd18+sVz9R+QRMjVAD5sdHSX6WCjVeS01DOTQ2GH0kNa9+ltUbV5O8nOTFJH9NVZacSU0VGCo1TkWIMS+GPg+Pj72Pnmx/ZypDnwMUn3xObhBm/97czoyqKCZp29hzsbb386wtOZ+AaRFqwPw5kqrSONIGpQfagPRgkofHBql7ZvDa/prkt0n+K9UP5FiScw7ZXDqU5N8z6tNyoL1/Dtk1Xbqc6g9yya6Y2ACMm5tWpcbmqNRYL7Oc/mHqCTBVQg2Y3wHMm23bmmoSOIQbz7aB6gupT96n5cMkryf5Y6oPyHsO01z7Tmqp4BdSocYuu6RrZ9t23q6Y2CCMmxuChtUpPNfWCDXWataNOoWEwFQJNWD+nU9NSTmaqo44mupTsSnVUHFa/TaOpUKWv0WgMe8OpPpmvJBarpX+DSsI6acxOYKNmxsa1U4jbBivCjlr19/1e3lphuGCnhqAX+DATZ1P8vtUsHAkFTRMy+lUY9BjDsPc25vRdBMWw7F2/uF+aBaGSo2tU3oufTXW/n42BQTwSxyYWxfbAOfD1Ce303ApFaicSYUbzLfVNvhYtSsWwrkk7yQ5blcwA0OgsSPTCzV2xZS5tRia3i7PwWsAmDgXG+jT2dRqI9NaZeRK2y5GOXAPLiT5ND7ZXxSvJPlHkrftiokOwGa5BOY825EKGHZketNPdqamu3D39/ezfD8PU1+cT8DULnpAf4aVEC5O+XqhnLUPH6Saur6S6VXzMBl/T6049EqqrwZM27ZMt1Ho9lgBZa2WMtueGrNuVApsMBqFQp8upqaDXExydUo3LsOnPq4b8+94qvfKMCf9e0ketFu683pq+eT/TAUbMAurGU1p2zKl59vatk2pKkHuPFSwAgmwYRicQJ/OZBRqTMPl9uiTl3681I7b+bZ9N8njdks33kiFGf9nkp9FL5tpDga51rDE6mqmE2psac+3rT23KY/9nkfuF4CpEGpAny6mpp9M+ybcnPO+/DYVaAzvk81JHrFb5t77qVDqF6lKDYHG5A2BrVDji1YyqpyYxvV/vDJkNUKNu3G1bbM8n5adT8C0CDWgT9OeLyvI6NdfUtNQdiW5NzUNZcVumVvHkvyybS8mOWUQNrVBmEq0G1vJqK/GNKxmukvIOp8mc8/gfAKmRqgBfZrFp4qXxzb68mqSR5McTq2IcsAumTsft+P0lyS/TvKrVE+NjWLWPQCGKjSfLH/Ramrlk3um9HzDkq47o1no3bo643NqUyoMM84ApsLFBvo9d6d5Az582jIs7Up/A+YTqWkMsyzlvpqaOnWhfb3cbnyHgO5ye39d3/x2vFJomqHa8pcMFm71KejlsXNmfOA+/H+XUr1xPk41dj2a5LXUaid/TPLmBhyEzUOlhlDji1Yz3UqN5fZ8Q08N+jufNBYHpj4wAvqjrJM7dSHJ55nuMsCft8H5iVSYcqU9/7k2qN+Uaz8hv5ovDzVuJ0xY7/NsfJBwo8dbDSiuXPdzDD/f+SSfppbc/TDJO0neSjUI/djblTmypW2rU3zO1bGN/ixlNAVl2tdtYAMSakCfrkz5JkF1Rv9WMt0g7FKq0eUrqT4RZ8f+/rNc2+j2+vDgSkaBRzK7AG/5Bq9rOPdudP5tGvu3S2NfX18KfikV7HySCjVOJPkgFTwxWwZf19qSUaPQafbiWcn0Vlth/S2NXds3xbRVYMKEGuDmm41haLw3rYHJkSQvp5pdHk1VJiy3m9szqYqNq2M3wFevCwSGm+FZNqm9PmwZgo3L+fJQI7nxlJlN7f+7nAo1rGrCvBvChZVMP9SY9nOy/mbdKwfYIIQaABtjYDKEGtP65POtVJXGS6klSoE+rx3DNJAtU37eoUqEfgk1gKkwJx/6NHxSbFoItztAGP/EddLOpaozjkSgQT/3Q7OuDJpHWzNqFLo85ecdNh/A3d09wlBVNsvzSagBTO2iA8BiG/+0dRqrCZxJ9Yo4btdD99eOIdiYpuF6ZQpKn4YphEMjaICJEmoAbJyBybQGJ+dTjUH1jID+rx3TCkNv9LxCjX7HF+MbwMQvOgBsjIHJtObFD0u3nrfr6YwGzDe+dkw71Ng6ds1yr3p31+BhCsqsDMu6qtQAJs4vCoCNca1fnuIN5qb4hBUW5doxBBvTtBIroCzC7xyVGsDULjpAn3yiyLwaprvssitwP9S1YWA67WBhfPUToUa/7xuVGoBf4sBNXRnbhBvczvtlWDFnGt3wV5LsTLLbrqcTSuW//HzeMuXn3JJRHyDHpN/zSaUGMBUuNNAvYQa3a5hbfaltkzZUaey16+loEMaN7xOXM5tqCdNP+j6fhveOcwuYyi8rABbb0Lhz2CZtNVWlsSfT/4QXWD9DoDHLUMO9qjEGgAsOLKBhKsG0qjVcK/p/vwyrkUwj1NicCjX2tw3mnU+Tb37tH6aCTNswHUilxp27mtG0w1m/f0wfAgxUAFizIdC4kOkts7o3yYG2AX0aqiW2zOi5NxsUA3ArQg2AxXchydkkn7XHadiX5FDbtjkE0KUh1Fia0XNvdq8KwK34RQHcrvEl2ujLxSRnknzaHqfhwSSHkzyR5DGHgA7uh6zU8EWzbNS5Go1C79YQQl2d8fm07JwCpnXRAWDxnU0FGp9N8TkPt+2Q3Q9dmmVPiyHQEKQD8KWEGgAbw9mx7eKUnvOhjEKNzQ4BdHmfOKtQY3NUagBwm7+sAFh8Z5N8kpqC8ukUn/fpJE+mAg6YV1Y/ubFZVkosR0+Nu3U5o1XSnE/AwvOLAmBjON22TzK9ZqFJ9dZ4LKagQI/mYfqJSg0AvpRyYOjTsP788Ohc5lZOJDme5GSmG2oko4ahv05yyaHIniT3pFaFWWrn8blU4PTxBr6mXfXWmDuzDBX01ADgthgIQd+DgCtTeq7xm0oDjz5dTPJeKtg4PeXnfirJ80n+muQ3G3T/785oidsH2p+3jh2b8+24nEzyfpK3U0HURrumzcqwUoOy+S+GCrMKNbZGpcZaXZ7huXw1ow9fACZKqAH93qhcdsPAHTreBsrTDjW2J/l6kheSfJTkyAba53tTVSqPp6bhPJpkX6pSY0u78b+Yqp4ZQo1jSY62/fS3VAXHRrimCUznyxAobJ7h8+upAcAtCTWgT0q1uRsn23Z6Bs/9dJLnUsHKx6lwY9E9k+QbbTucCjYeTnJ/RlUag88z6nsyhBr/TPUk+XOSN7x9mbJNGVWwzMLSjJ8fgE4INYDbcdkuWAhXMuqtcf4GA+tJeig1BeWjVOXBr5OcWeB9/Z0kP0pVpzyXCjV2fcm/35KalvJAkq9mVNGyPxWCrCb5u7cwUzQPU3KGZV25M1cz3SmqN3p+H7wAU/1lAcDGcTIVbJxIVQ1M0zNJTqVCjc+SvLig+/i5VKDxk1S48fhdfI/727Y7NX3nUtt373sLMyWzrtQYfw0AcFNCDYCN5USSD9o27VDjnlTlwpm2fZzFqz54Ksn3kny/PT66Dt9vJdVz4/0kH0blFNOx3LZZhgoahQJwS0INgI3lgyTvpno2fHsGz7+/DfbPpqo1PmmvZxE8llGFxg+z9kBj8JX2fd9rx++P3sZMwaxXP7n+NVx0SG7bsPKIABTYEJT0AWy8m92hEeXHM3oNjyb5VipU+V6qj0Tv9qaChx+nqjQeX+fv/81UT5JvJDngbcwUDP0s5iXUAICb/sICYGP5IBVqvJkKF2bhm6lqjQupZnb/ner30aN7UoHGvyf5tyRPTuh5nk01D30nFUzBJA1TT2Z5ryjUuDuadAIbilAD+nXFLuAuHUsFGq+mpjbsntHr+H4brGxLre7xs1TPiJ7cm6rO+Gnbnpzgcz2Varb6ZmqJ138t2PXs8hw8v8HgtfeIK5lto9Dl9hpWHY47fj/PcvUTgKn/wgJgY7nQBsT/SC01+r0ZvY6l1BSU1dSSpstJfpmqIunBnlTvjJ+07ckpPOehth3MYoUazJ9h5ZNZhhoqNe6OcA7YUIQaABvTW6lKjceTPJ2aQjErX08tWzpUbbyU5K9zvv8Opabu/CBVqfGdKT3v/rbtbYNNjQCZlGH6yTysfqJSA4CbEmpA36bZ3fxylLMukpNJ/pbkkdSUhu/O+PV8JcnWVLiyp22vJDk9Z/ttNdWw87n2+J3UMrXTsj3JzrZtT/LpAr0nl5yWc2Wo1JhlqLGlXRe2OBwA3IxQA2Djei21DOmRzD7USJKHUs02t6f6fOxPVZO8lvlYzvGrqaqWF1KrkDyb5GszeB3z8Ak6i0+lBgBdEGoAbFwXkryeCjVOpKY0zNreVLBxTyrUeLRtb6Z6SJydwWt6MMkTqUDj2VSVxrOZzVK0F5N83h5VTTFJemoA0AWhBvRrmp36DZ4W12sZVUPsnZPXtD21ROrBFiZ8NbXaxz/ba30ryfkpvI57U80/v5paeeTptj0zw9+fn6am5HyS2QQ8bBybMj+VGlsdjjv+nT3N6akAMyXUAHDzeyRVsfF8CxTmxVClcSgVZBxpXx9pfz6a9e8psZQKdx5OBSqH2/Zke3xoxvvkRNtOGrAwYfNQqbEpKjUAuAWhBgBvpSog/pr56K1xvcNj26HUVJSjSd5tA/yPk5xKVS+cTvLZHXzvrakQ495Uc9K9GYUaw/MdSoUr8+BY20542zJh89BTYzlCDQBuQagBwLEkv09N99iVmmIxjx5r2ztJjrftZCrU+Ghs+zgVcFxMcik1TWtYWWOp/e5bSbIj1btjX5L7U2HGvva4v+2Pe+bo538vFea83Y7ZornqVJwry9FTo1dXMtvVypZjNSNgioQaACS1fOqBJPelKhb2z/FrfbhtSXIuVZ1xItdOzfgk1Xfj+oaaQzn71tSyqMMqK0OFxr45/t34ZmrqzZvRT4PpDExnXakxPL/7VQBuyi8JAJL65P+lVKBxIPMdaozb1rYH258/S1VrjIcal9p/G6/S2J5Rpca9HfycF1N9T15NBRuLSKXGfBl6asx6+sl4uKJpNQBfINSAvgcA0x4EKCddbK+mKhceTgUbT3f4M+xo2yIem78k+VsWs0pjHgKNJYPmubvuL+XahqWOz+2x+gmwoWyyC6DrQcC058z6JHXxvdK2V2OKw7z4JMmfU6HGmws8cJ51aOr6dq2hQmLJ+6P7ewWAhSbUgH5dTpXVX5zS8w2f/PikbLGdbYPnV9pAmtk6l+TXSX7TjssiD6Bn2ZDSte3G94iznn4yVCQ6Nne332YZaAiigKkx/QT6Nq3y0qupAOXzKGfdCF5J8lCqkebmJN+xS2biTJKfJ/lZkl+klrBdVLMONYZrnE+1RzZl9o1CL7VtCPG5u/f20gZ87wAbiFADuB3Dp5izXCKO6fk8yZ9SzTRX2u+K5+2WqTqfqtD477b9foMMoGc58Et8snz9MZn10pxDtYFKmrt7P8/DOS3YACZOqAF933CutG3SLuXaYIPFd3Ts/bU1yf1JHrFbpuJykt+mVqP5dZKX7ZKJUyr/RbOeejLwe2dt++3qjN9Dyw4FMGlCDejbMOCctM8zKs1Wnr1xHMkoPNudCja22y0TdTHJr5K8mJpy8psNMqCb9eBrGIAJNq49JvPgcqbXO2qRjt0wPfXKDIIFIRQwVUIN6NfVTK+082Iq1JiHgQfT9XoqONudZE+S/8h0qoM2orOpqoxfpEKNl7JxVqCZ9fKTwyfKQo1rB6ZX5uR9IdQA4KaEGtCvi2Pb5ik813jDNjaWV5KsZlSl8aOo2FhvFzIKNH6eCjROb6Cf/3JmG5gO/SMY+XwOwoSL7XWcdzgAuBmhBvTrs7adTbJtws91vj3XuajU2Kh+2wZ+F9p77j9S1Rusz4D+V6lA479SocZG+2T60ox/5vFSfcrZJJ+26//WGb6Gz1IrAXHn7+dZhYVXcu0HLwATJdSAvm84z7TH+yf8XBdSgcZZu31D+017L1xMfar9kwg21upiqkLjxVSw8dIGHQRczGyrwC7OwWuYN+fa+T7LUONMKlj51OG441BhliuWDVOGzqcqbQAmSqgB/TrbtgtTei43liS11OvFjPpq/DDJfXbLXfkkye9SQcYv2uNG/UR61p/ons8osKMMgca5JPfO6DWcHtu48/PpYmYTagyBxnCfAjBRQg3o1/AJ1qRvGC62G8pTbYO/pSo1Lrcb1+8nedhuuSNHU5Uvv0tVamzkQCNjg5+rmU2zzlPtOmcANjJM+zid5MCMXsOJtvndc2fGA4WzmX4PpM/a/cnpVCgGMFFCDejXmUznU5Dz7cbkZNsgqeah4z02XkjyrN1yS1dS1S5/SPLr9vhKplNxNc/Ojg2g753Bcw+DZ70brv0dczrJh0mensHzf9C2Ew7FXR27T8e2vVN+/k/a/YIwCpgKoQb06/OxG5dJ3/B/PHbTD4M/ZTQ1aajkeSFWRrmZc6nKjN+lqjR+neQNuyVp76GT7Rpz75Sf+3iSY20Azcjptk+Ophq5Tvue8Vjb/N65u9/bp8a2x6f8/CfdMwDTJNSAvg03DpMs2T7Vbmzfi5UB+KLXM5qedKJ9/VxMR7nee6mKjGHKyW/agI1ypoULbyf56pSf+19J3kryrsPwBW8l+Xt7735ris97OsmR9vzOk7tzLMk77XGazV4vtevdu44dMC1CDejbiTYQeCfJIxN+Djcn3Mzx1FKkQ0+C00meT/KMXZPPkrzatldSS+P+MclHds0XvNMGss9les1n30lVy7yZqkjgWm8m+WsqaNqf5KEpPe9rqcD0zXZ94c4da8HCEGw8PuXnFRICUyPUgL592G5Y3spkQo1LGZUAv2938yXOpJYkHXqwnGjvzyemOBCaJ+faAP1IqrHqa0n+ktHqMXzR0VT487fUcsHTcKQdmyPRJPRGPmvH49Ek+5I8kNHKR5McFP+tvReOOAR37XQqFHqrnVuPT/E8fjs+CAGmaHllZcVegL5vOB9Ick+S3UnuX+fv/1qSF9tg1Q0Kt3I1VXZ8ItWHZWgWdzoVou/eIPvhjXbevJha1eTXqekmr8UUri/zaZJdqZ4sO1OVAZP0z3Z8ftUeLzsEN/RRO3+3tOMyyallF1PLGw9LHL9u96/5HuHeVOXTgXZ+Tfra99t2zfut3Q9Mi1AD+nal3Wxua9t97aZzPXzebvZfbAMyn2Jyu06nPq07nQo3TqUCjs/agHVRG4l+lFrNZGgC+nKqf8afYlWN23W2Xcu2tsHYvRMc7L04dn0zxeHW5/SWdmz2T/C4/D7Jf7ffPX+IEHA9fo8PgfLuJF+Z4HO9nwoyfpMKpDQJBaZGqAH9+7DdbK62weLhdfq+f0x9Wvbr1CfMcCeupBrMvpvR0n4ft+3i2Pt1Uc7BV5P8st3Q/6o9/sGN/R070wZiK+26tjtVibaeLrbr2i9Tocaf7PZbupCqpBlC9L1Z/0/9X2vH5L8iSF9Pp5PsSAWFq5nMdMB3UoHGrzJaqhpgaoQasBg+aIOAbanpKGtdk/71dmP583Zz+bldzBoGQ++n5lh/0L4+3h4/aO+t5fQVcFxtN/H/yCjE+OXYIPnl1HQt0xnuzok2oB1WdNqZZM86DvB+mdEUh99Fj5Pbdabtv03tHNie9Wvo+ko7j37ejo9GuuvncqpS7kqqT9blVLCxvE7f/0iqMuO/2/F7yS4Hpk2oAYt107I19enZ4dSnnHfjWJL/bDcnL7q5ZJ1cSIUY/0pVNgyhxkepSo7P2vt3nsONS6mme79LfSr5ckZVGb9pf36//TvW5sNU4HW1Xd+G6oC1eC9fDJ8+s6vvyLDC0eVUGLQpFaRvusvv92k7n37RjsuvUoEh6+uzVJXc5xmFeI+u4bgNXs1ousmvoo8GMCNCDVgcZ1MrLqy0bX/ufF36D5P8f5L8f9sNyr/sViYQDHyYatL4YSrU+LgNls609/DmVLn0PA0IjqSmKQw9M37XBsW/TfL3VDBj/v/6+mBsEHaxDaTvzZ0HtpdSy5IOTY9/LdBYkxPtPP287dsLqdDpTvs5/asdixdTn/K/KNCY+HVsCAuHwHBPO3Z36ly7Fr7crodDyAsg1ADWbOhbcLENEC+1YONWNy2ftYHZz5L8v1OVGgINJm1YcvBUCwU+SlU6vNsGNx+mKpAupz5R3DKF13Qp9enx8dQ0rD+O3bD/vn39u3ZD/2a7uWeywcbQZPbj9h45n5qaspqbl9BfaYPvN9rA+VcZNT425WR9go3T7bh80o7NmTZQ3nGT43KpnetvJ/lzOxa/zqhK42O7deIupJo4f9aO14n29aZ2Pm2+xf//ebvuDSs6DQHv7+1aYJaWtm/fbi/A4nk0yfNJnklNRTmQ0cooK+2G/3yquuNMGzi80QZwv4zmhkzf5iQHUxVGe1OfIA4d++9rfx4eh2U/h0Htpva+Hr4eHm800B3K5i9cdw6cve7PJ1uwcaJ9faJtx9sgjuna265ljyd5LMmD7e92t/fB8AnNcGw/bQPoYcrTG6nqIEtTr697rzsuB9q2t52jy9edV6fbefT+2DF51W6ciUNJnkry1dSqKAfGzqmVdg29klGl1Pl2Tr2TCnxfa8fvqF0JCDWASVlN8mQLOB5qA8Jd7e8vtxv/M6lPxz5I9Qr4ezQFZT7sbO/Xe8YCjd1jocbOVAXSEGwMj5vb4/J1YcbQm+HS2A36uYw+af5s7Jz4dCzU+CimKcyT3W3wtW/sPTGs6nClXb8uZFQ9cCKjFXiY7Pn6cKq/xr4k97fzc/MNzrWT7XfOO9F/Zh4caAHHEGrcm1HVxpWx6+XZFmoca0GGMAMQagDTO89Tn2re3248hz4b58YGb+/bTXRge3sP77wu2NiSUS+Z4RPGIdS42m7Mh5vzyxnNKb8+1Bhu3AV7fdicCr22ZlSBdrEd10/asWf6hqV4h0qN4Vw7F8u0zvu9wv2pUGNbRtVP58fOqY9jVSdgHi9gQg3YcIayfE0NWZT38zCovVGocTnXTjsZL6fWVwEAoHOb7QLYcIQZLNr72ae/AAAb1Ca7AAAAAOiRUAMAAADoklADAAAA6JJQAwAAAOiSUAMAAADoklADAAAA6JJQAwAAAOiSUAMAAADoklADAAAA6JJQAwAAAOiSUAMAAADoklADAAAA6JJQAwAAAOiSUAMAAADoklADAAAA6JJQAwAAAOiSUAMAAADoklADAAAA6JJQAwAAAOiSUAMAAADoklADAAAA6JJQAwAAAOiSUAMAAADoklADAAAA6JJQAwAAAOiSUAMAAADoklADAAAA6JJQAwAAAOiSUAMAAADoklADAAAA6JJQAwAAAOiSUAMAAADoklADAAAA6JJQAwAAAOiSUAMAAADoklADAAAA6JJQAwAAAOiSUAMAAADoklADAAAA6JJQAwAAAOiSUAMAAADoklADAAAA6JJQAwAAAOiSUAMAAADoklADAAAA6JJQAwAAAOiSUAMAAADoklADAAAA6JJQAwAAAOiSUAMAAADoklADAAAA6JJQAwAAAOiSUAMAAADoklADAAAA6JJQAwAAAOiSUAMAAADoklADAAAA6JJQAwAAAOiSUAMAAADoklADAAAA6JJQAwAAAOiSUAMAAADoklADAAAA6JJQAwAAAOiSUAMAAADoklADAAAA6JJQAwAAAOiSUAMAAADoklADAAAA6JJQAwAAAOiSUAMAAADoklADAAAA6JJQAwAAAOiSUAMAAADoklADAAAA6JJQAwAAAOiSUAMAAADoklADAAAA6JJQAwAAAOiSUAMAAADoklADAAAA6JJQAwAAAOiSUAMAAADoklADAAAA6JJQAwAAAOiSUAMAAADoklADAAAA6JJQAwAAAOiSUAMAAADoklADAAAA6JJQAwAAAOiSUAMAAADoklADAAAA6JJQAwAAAOiSUAMAAADoklADAAAA6JJQAwAAAOiSUAMAAADoklADAAAA6JJQAwAAAOiSUAMAAADoklADAAAA6JJQAwAAAOiSUAMAAADoklADAAAA6JJQAwAAAOiSUAMAAADoklADAAAA6JJQAwAAAOiSUAMAAADoklADAAAA6JJQAwAAAOiSUAMAAADoklADAAAA6JJQAwAAAOiSUAMAAADoklADAAAA6JJQAwAAAOiSUAMAAADoklADAAAA6JJQAwAAAOiSUAMAAADoklADAAAA6JJQAwAAAOiSUAMAAADoklADAAAA6JJQAwAAAOiSUAMAAADoklADAAAA6JJQAwAAAOiSUAMAAADoklADAAAA6JJQAwAAAOiSUAMAAADoklADAAAA6JJQAwAAAOiSUAMAAADoklADAAAA6JJQAwAAAOiSUAMAAADoklADAAAA6JJQAwAAAOiSUAMAAADoklADAAAA6JJQAwAAAOiSUAMAAADoklADAAAA6JJQAwAAAOiSUAMAAADoklADAAAA6JJQAwAAAOiSUAMAAADoklADAAAA6JJQAwAAAOiSUAMAAADoklADAAAA6JJQAwAAAOiSUAMAAADoklADAAAA6JJQAwAAAOiSUAMAAADoklADAAAA6JJQAwAAAOiSUAMAAADoklADAAAA6JJQAwAAAOiSUAMAAADoklADAAAA6JJQAwAAAOiSUAMAAADoklADAAAA6JJQAwAAAOiSUAMAAADoklADAAAA6JJQAwAAAOiSUAMAAADoklADAAAA6JJQAwAAAOiSUAMAAADoklADAAAA6JJQAwAAAOiSUAMAAADoklADAAAA6JJQAwAAAOiSUAMAAADoklADAAAA6JJQAwAAAOiSUAMAAADoklADAAAA6JJQAwAAAOiSUAMAAADoklADAAAA6NL/H2m3vrg5H2g3AAAAAElFTkSuQmCC";

const fbApp = initializeApp(FIREBASE_CONFIG);
const db = getDatabase(fbApp);

function pad(n) { return String(n).padStart(2, "0"); }
function todayStr() {
  const d = new Date();
  const yy = String(d.getFullYear()).slice(2);
  return yy + "." + pad(d.getMonth()+1) + "." + pad(d.getDate()) + " " + pad(d.getHours()) + ":" + pad(d.getMinutes());
}
function nowStr() {
  const d = new Date();
  const yy = String(d.getFullYear()).slice(2);
  return yy+"."+pad(d.getMonth()+1)+"."+pad(d.getDate())+" "+pad(d.getHours())+":"+pad(d.getMinutes())+":"+pad(d.getSeconds());
}

function toDateObj(s) {
  if (!s) return null;
  // "25.06.01 14:30" 또는 "2025.06.01" 형태 모두 처리
  const dateOnly = s.split(" ")[0]; // 시간 제거
  const p = dateOnly.split(".");
  if (p.length < 3) return null;
  const year = p[0].length === 2 ? 2000 + parseInt(p[0]) : parseInt(p[0]);
  return new Date(year, parseInt(p[1])-1, parseInt(p[2]));
}
function fmt(n) { return Number(n).toLocaleString("ko-KR"); }

export default function App() {
  const [role, setRole] = useState(null);
  const [userName, setUserName] = useState("");
  const [pw, setPw] = useState("");
  const [pwErr, setPwErr] = useState(false);
  const [jobs, setJobs] = useState([]);
  const [syncMsg, setSyncMsg] = useState("");
  const [tab, setTab] = useState("all");
  const [page, setPage] = useState(1);
  const PAGE_SIZE = 100;

  // form
  const [item, setItem] = useState("");
  const [worker, setWorker] = useState("");
  const [qty, setQty] = useState("");
  const [price, setPrice] = useState("");
  const [formErr, setFormErr] = useState(false);

  // search
  const now = new Date();
  const [srchFrom, setSrchFrom] = useState(now.getFullYear()+"-"+pad(now.getMonth()+1)+"-01");
  const [srchTo, setSrchTo]     = useState(now.getFullYear()+"-"+pad(now.getMonth()+1)+"-"+pad(now.getDate()));
  const [srchWorker, setSrchWorker] = useState("");
  const [report, setReport] = useState(null);

  // modal
  const [modal, setModal] = useState(null);
  const [modalVal, setModalVal] = useState("");
  const modalRef = useRef(null);
  const fileRef  = useRef(null);
  const [uploadMsg, setUploadMsg] = useState("");
  const [showLog, setShowLog] = useState(false);
  const [logTab, setLogTab] = useState("activity"); // activity | access
  const [accessLogs, setAccessLogs] = useState([]);
  const [activityLogs, setActivityLogs] = useState([]);
  const refItem   = useRef(null);
  const refWorker = useRef(null);
  const refQty    = useRef(null);
  const refPrice  = useRef(null);
  const refSrchFrom   = useRef(null);
  const refSrchTo     = useRef(null);
  const refSrchWorker = useRef(null);

  // ── 활동 로그 저장 ────────────────────────────────────
  const logActivity = async (action, detail) => {
    try {
      await push(ref(db, "activity_logs"), {
        name: userName,
        role: role,
        action,
        detail,
        time: nowStr(),
        ts: Date.now()
      });
    } catch(e) {}
  };

  // ── 접속 이력 불러오기 (정태식만) ───────────────────────
  const loadAccessLogs = async () => {
    try {
      const logRef = ref(db, "access_logs");
      onValue(logRef, (snapshot) => {
        const data = snapshot.val();
        if (data) {
          const arr = Object.entries(data)
            .map(([k, v]) => ({ ...v, key: k }))
            .sort((a, b) => b.ts - a.ts)
            .slice(0, 100);
          setAccessLogs(arr);
        } else { setAccessLogs([]); }
      }, { onlyOnce: false });

      const actRef = ref(db, "activity_logs");
      onValue(actRef, (snapshot) => {
        const data = snapshot.val();
        if (data) {
          const arr = Object.entries(data)
            .map(([k, v]) => ({ ...v, key: k }))
            .sort((a, b) => b.ts - a.ts)
            .slice(0, 200);
          setActivityLogs(arr);
        } else { setActivityLogs([]); }
      }, { onlyOnce: false });
    } catch(e) {}
  };

  // ── Firebase 실시간 리스닝 ─────────────────────────────
  useEffect(() => {
    if (!role) return;
    if (userName === "정태식") loadAccessLogs();
    const jobsRef = ref(db, "jobs");
    const unsub = onValue(jobsRef, (snapshot) => {
      const data = snapshot.val();
      if (data) {
        const arr = Object.entries(data)
          .map(([fbKey, v]) => ({ ...v, fbKey }))
          .sort((a, b) => b.createdAt - a.createdAt);
        setJobs(arr);
      } else {
        setJobs([]);
      }
      setSyncMsg("✅ 동기화됨");
      setTimeout(() => setSyncMsg(""), 2000);
    });
    return () => unsub();
  }, [role]);

  // ── 로그인 ────────────────────────────────────────────
  const doLogin = async () => {
    const user = USERS[pw.trim()];
    if (!user) { setPwErr(true); setPw(""); return; }
    setRole(user.role);
    setUserName(user.name);
    setPwErr(false); setPw("");
    // 접속 이력 Firebase에 저장
    try {
      const logRef = ref(db, "access_logs");
      await push(logRef, {
        name: user.name,
        role: user.role,
        time: nowStr(),
        ts: Date.now()
      });
    } catch(e) {}
  };

  // ── 등록 ─────────────────────────────────────────────
  const addJob = async () => {
    const q = parseInt(qty), p = parseInt(price) || 0;
    if (!item.trim() || !worker.trim() || !q || q < 1) { setFormErr(true); return; }
    setFormErr(false);
    const jobsRef = ref(db, "jobs");
    await push(jobsRef, {
      item: item.trim(), worker: worker.trim(),
      qty: q, price: p,
      date: todayStr(), status: "pending", doneDate: null,
      createdBy: userName,
      createdAt: Date.now()
    });
    setItem(""); setWorker(""); setQty(""); setPrice("");
    await logActivity("등록", item.trim()+" / "+worker.trim()+" / 수량:"+q+(p?" / 단가:"+fmt(p):""));
    setSyncMsg("💾 저장됨"); setTimeout(() => setSyncMsg(""), 2000);
  };

  // ── 엑셀 업로드 ──────────────────────────────────────
  const handleExcel = async (e) => {
    const file = e.target.files[0];
    if (!file) return;
    fileRef.current.value = "";

    const reader = new FileReader();
    reader.onload = async (evt) => {
      try {
        const wb = XLSX.read(evt.target.result, { type: "array", cellDates: true });
        const ws = wb.Sheets[wb.SheetNames[0]];
        const rows = XLSX.utils.sheet_to_json(ws, { header: 1, defval: "" });

        // 헤더 행 찾기 (일자/품목/작업자/수량/단가 포함된 행)
        let dataStart = 0;
        for (let i = 0; i < Math.min(rows.length, 5); i++) {
          const row = rows[i].map(c => String(c).replace(/\s/g,"").toLowerCase());
          if (row.some(c => c.includes("품목") || c.includes("작업자") || c.includes("수량"))) {
            dataStart = i + 1;
            break;
          }
        }

        // 컬럼 인덱스 자동 감지
        const headerRow = rows[dataStart - 1] || [];
        const colIdx = { date:-1, item:-1, worker:-1, qty:-1, price:-1 };
        headerRow.forEach((h, i) => {
          const s = String(h).replace(/\s/g,"").toLowerCase();
          if (s.includes("일자") || s.includes("날짜") || s.includes("date")) colIdx.date = i;
          else if (s.includes("품목") || s.includes("item"))                   colIdx.item = i;
          else if (s.includes("작업자") || s.includes("worker"))               colIdx.worker = i;
          else if (s.includes("수량") || s.includes("qty"))                    colIdx.qty = i;
          else if (s.includes("단가") || s.includes("price"))                  colIdx.price = i;
        });

        // 헤더 못 찾으면 순서대로 (일자,품목,작업자,수량,단가)
        if (colIdx.item === -1) { colIdx.date=0; colIdx.item=1; colIdx.worker=2; colIdx.qty=3; colIdx.price=4; dataStart=1; }

        const dataRows = rows.slice(dataStart).filter(r => r.some(c => c !== ""));
        if (!dataRows.length) { setUploadMsg("❌ 데이터가 없습니다."); return; }

        let successCnt = 0;
        const jobsRef = ref(db, "jobs");

        for (const row of dataRows) {
          const rawDate = row[colIdx.date];
          const itemVal = String(row[colIdx.item] || "").trim();
          const workerVal = String(row[colIdx.worker] || "").trim();
          const qtyVal = parseInt(row[colIdx.qty]) || 0;
          const priceVal = parseInt(row[colIdx.price]) || 0;

          if (!itemVal || !workerVal || qtyVal < 1) continue;

          // 날짜 파싱
          let dateStr = todayStr();
          if (rawDate) {
            if (rawDate instanceof Date) {
              dateStr = rawDate.getFullYear()+"."+pad(rawDate.getMonth()+1)+"."+pad(rawDate.getDate());
            } else {
              const s = String(rawDate).replace(/[.\-/]/g, ".");
              const parts = s.split(".");
              if (parts.length >= 3) {
                dateStr = parts[0]+"."+pad(parseInt(parts[1]))+"."+pad(parseInt(parts[2]));
              }
            }
          }

          await push(jobsRef, {
            item: itemVal, worker: workerVal,
            qty: qtyVal, price: priceVal,
            date: dateStr, status: "pending", doneDate: null,
            createdBy: userName,
            createdAt: Date.now()
          });
          successCnt++;
        }

        await logActivity("엑셀업로드", successCnt+"건 업로드 ("+file.name+")");
        setUploadMsg("✅ " + successCnt + "건 업로드 완료!");
        setTimeout(() => setUploadMsg(""), 4000);
      } catch(err) {
        setUploadMsg("❌ 파일 읽기 오류: " + err.message);
        setTimeout(() => setUploadMsg(""), 4000);
      }
    };
    reader.readAsArrayBuffer(file);
  };

  // ── 상태 변경 (마스터) ────────────────────────────────
  const setStatus = async (j, s) => {
    if (role !== "master") return;
    await update(ref(db, "jobs/" + j.fbKey), {
      status: s,
      doneDate: s === "done" ? todayStr() : null
    });
    const label = s==="done"?"완료처리" : s==="partial"?"일부완료처리" : "미완료복귀";
    await logActivity(label, j.item+" / "+j.worker);
  };

  // ── 삭제 (마스터) ─────────────────────────────────────
  const deleteJob = async (j) => {
    if (role !== "master") return;
    if (!confirm("삭제하시겠습니까?")) return;
    await logActivity("삭제", j.item+" / "+j.worker+" / 수량:"+j.qty);
    await remove(ref(db, "jobs/" + j.fbKey));
  };

  // ── 인라인 수정 모달 ──────────────────────────────────
  const canEdit = (j) => role === "master" || j.status !== "done";
  const openModal = (j, field) => {
    if (!canEdit(j)) return;
    setModal({ fbKey: j.fbKey, field, sub: j.item + " / " + j.worker });
    const initVal = field === "qty" ? j.qty : field === "price" ? j.price : field === "item" ? j.item : j.worker;
    setModalVal(String(initVal ?? ""));
    setTimeout(() => modalRef.current?.focus(), 80);
  };
  const saveModal = async () => {
    if (modal.field === "item" || modal.field === "worker") {
      if (!modalVal.trim()) return;
      await update(ref(db, "jobs/" + modal.fbKey), { [modal.field]: modalVal.trim() });
      await logActivity("수정", modal.field==="item"?"품목→"+modalVal.trim():"작업자→"+modalVal.trim()+" ("+modal.sub+")");
      setModal(null);
      return;
    }
    const v = parseInt(modalVal);
    if (isNaN(v) || (modal.field === "qty" && v < 1) || v < 0) return;
    await update(ref(db, "jobs/" + modal.fbKey), { [modal.field]: v });
    await logActivity("수정", (modal.field==="qty"?"수량":"단가")+"→"+fmt(v)+" ("+modal.sub+")");
    setModal(null);
  };

  // ── 검색 ─────────────────────────────────────────────
  const doSearch = () => {
    if (!srchFrom || !srchTo) { alert("날짜를 선택해 주세요."); return; }
    const from = new Date(srchFrom), to = new Date(srchTo); to.setHours(23,59,59);
    const wf = srchWorker.trim().toLowerCase();
    const inRange = jobs.filter(j => {
      const d = toDateObj(j.date);
      if (!d || d < from || d > to) return false;
      const wRaw3 = String(j.worker||'').trim();
      const wKey3 = wRaw3.startsWith('원광') ? '원광' : wRaw3.slice(0,3);
      return !wf || wKey3.toLowerCase().includes(wf);
    });
    if (!inRange.length) { alert("해당 기간에 작업 내역이 없습니다."); return; }
    const map = {};
    inRange.forEach(j => {
      const wRaw = String(j.worker||'').trim();
      const workerKey = wRaw.startsWith('원광') ? '원광' : wRaw.slice(0,3);
      if (!map[workerKey]) map[workerKey] = {
        totalCnt:0, totalQty:0, totalAmount:0,
        doneCnt:0,  doneQty:0,  doneAmount:0,
        etcCnt:0,   etcQty:0,   etcAmount:0
      };
      const q = Number(j.qty), p = Number(j.price);
      map[workerKey].totalCnt++;
      map[workerKey].totalQty += q;
      map[workerKey].totalAmount += q * p;
      if (j.status === "done") {
        map[workerKey].doneCnt++;
        map[workerKey].doneQty += q;
        map[workerKey].doneAmount += q * p;
      } else {
        map[workerKey].etcCnt++;
        map[workerKey].etcQty += q;
        map[workerKey].etcAmount += q * p;
      }
    });
    setReport({ from: srchFrom.replaceAll("-","."), to: srchTo.replaceAll("-","."), total: inRange.length, map });
  };

  // ── 필터 ─────────────────────────────────────────────
  const pending = jobs.filter(j => j.status === "pending");
  const partial = jobs.filter(j => j.status === "partial");
  const done    = jobs.filter(j => j.status === "done");
  const filtered = tab==="pending"?pending : tab==="partial"?partial : tab==="done"?done : jobs;
  const totalPages = Math.max(1, Math.ceil(filtered.length / PAGE_SIZE));
  const paginated = filtered.slice((page-1)*PAGE_SIZE, page*PAGE_SIZE);

  const rowCls = j => j.status==="done"?"row-d":j.status==="partial"?"row-pt":"row-p";
  const txCls  = j => j.status==="done"?"tx-d":j.status==="partial"?"tx-pt":"tx-n";

  const Pill = ({j}) => j.status==="done"
    ? <span className="pill pill-g">완료</span>
    : j.status==="partial"
      ? <span className="pill pill-y">일부완료</span>
      : <span className="pill pill-r">지급</span>;

  const ActBtns = ({j}) => {
    if (role !== "master") return <span className="lock"><i className="ti ti-lock"/></span>;
    if (j.status === "pending") return <div className="act-group">
      <button className="act-btn act-pt" onClick={()=>setStatus(j,"partial")} style={{fontSize:8,padding:"3px 6px"}}><i className="ti ti-triangle" style={{fontSize:7}}/> 일부완료</button>
      <button className="act-btn act-d"  onClick={()=>setStatus(j,"done")} style={{fontSize:14,padding:"8px 14px",fontWeight:600}}><i className="ti ti-check"/> 완료</button>
    </div>;
    if (j.status === "partial") return <div className="act-group">
      <button className="act-btn act-d" onClick={()=>setStatus(j,"done")} style={{fontSize:14,padding:"8px 14px",fontWeight:600}}><i className="ti ti-check"/> 완료</button>
      <button className="act-btn act-u" onClick={()=>setStatus(j,"pending")} style={{fontSize:8,padding:"3px 6px"}}><i className="ti ti-arrow-back-up"/> 취소</button>
    </div>;
    return <div className="act-group">
      <button className="act-btn act-pt" onClick={()=>setStatus(j,"partial")} style={{fontSize:8,padding:"3px 6px"}}><i className="ti ti-triangle" style={{fontSize:7}}/> 일부완료</button>
      <button className="act-btn act-u"  onClick={()=>setStatus(j,"pending")} style={{fontSize:8,padding:"3px 6px"}}><i className="ti ti-arrow-back-up"/> 미완료</button>
    </div>;
  };

  const reportRows = report ? Object.entries(report.map).sort((a,b)=>a[0].localeCompare(b[0],"ko")) : [];

  // ── 로그인 화면 ───────────────────────────────────────
  if (!role) return (
    <div style={S.loginOverlay}>
      <div style={S.loginBox}>
        <img src={LOGO} alt="Petit" style={S.loginLogo}/>
        <h2 style={S.loginTitle}>하청 작업 관리</h2>
        <label style={S.lbl}>비밀번호</label>
        <input style={{width:"100%",height:52,border:"1px solid #ccc",borderRadius:10,padding:"0 16px",fontSize:18,fontFamily:"inherit",outline:"none",background:"#fafaf8",marginBottom:20,boxSizing:"border-box"}} type="password" placeholder="비밀번호 입력"
          value={pw} onChange={e=>{setPw(e.target.value);setPwErr(false);}}
          onKeyDown={e=>e.key==="Enter"&&doLogin()}/>
        <button style={S.loginBtn} onClick={doLogin}>로그인</button>
        {pwErr && <p style={S.loginErr}>비밀번호가 올바르지 않습니다.</p>}
      </div>
    </div>
  );

  // ── 메인 앱 ───────────────────────────────────────────
  return (
    <div style={S.wrap}>
      {/* 헤더 */}
      <div style={S.header}>
        <div style={S.headerLeft}>
          <img src={LOGO} alt="Petit" style={S.headerLogo}/>
          <div>
            <h1 style={S.h1}>하청 작업 관리</h1>
            <p style={S.sub}>부자재 지급 → 일부완료 → 제품 완성 수거</p>
          </div>
        </div>
        <div style={S.headerRight}>
          <span style={role==="master"?S.chipMaster:S.chipStaff}>
            {role==="master"?"👑 "+userName : "👤 "+userName}
          </span>
          {userName === "정태식" && (
            <button style={{...S.logoutBtn, color:"#1a56db", borderColor:"#1a56db"}}
              onClick={()=>setShowLog(true)}>
              👁 접속현황
            </button>
          )}
          <button style={S.logoutBtn} onClick={()=>{setRole(null);setUserName("");setJobs([]);setShowLog(false);}}>로그아웃</button>
        </div>
      </div>

      {/* 배지 */}
      <div style={S.badges}>
        <span style={S.badgeRed}>미완료 {pending.length}건</span>
        <span style={S.badgeYellow}>일부완료 {partial.length}건</span>
        <span style={S.badgeGreen}>완료 {done.length}건</span>
        <span style={S.badgeBlue}>🔴 실시간 공유</span>
      </div>
      {syncMsg && <p style={S.syncBar}>{syncMsg}</p>}

      {/* 등록 폼 */}
      <div style={S.card}>
        <p style={S.secTitle}>➕ 신규 작업 등록 <span style={S.todayChip}>등록시각: {todayStr()}</span></p>
        <div style={S.formGrid}>
          <div><label style={S.lbl}>품목</label><input ref={refItem} style={S.inp} type="text" placeholder="품목명" value={item} onChange={e=>setItem(e.target.value)} onKeyDown={e=>e.key==="Enter"&&refWorker.current?.focus()}/></div>
          <div><label style={S.lbl}>작업자</label><input ref={refWorker} style={S.inp} type="text" placeholder="작업자명" value={worker} onChange={e=>setWorker(e.target.value)} onKeyDown={e=>e.key==="Enter"&&refQty.current?.focus()}/></div>
          <div><label style={S.lbl}>수량</label><input ref={refQty} style={S.inp} type="number" min="1" placeholder="수량" inputMode="numeric" value={qty} onChange={e=>setQty(e.target.value)} onKeyDown={e=>e.key==="Enter"&&refPrice.current?.focus()}/></div>
          <div><label style={S.lbl}>단가(원)</label><input ref={refPrice} style={S.inp} type="number" min="0" placeholder="단가" inputMode="numeric" value={price} onChange={e=>setPrice(e.target.value)} onKeyDown={e=>e.key==="Enter"&&addJob()}/></div>
          <div style={{display:"flex",alignItems:"flex-end",gap:8,gridColumn:"1/-1"}}>
            <button style={S.btnAdd} onClick={addJob}>+ 등록</button>
            <button style={S.btnExcel} onClick={()=>fileRef.current.click()}>
              📂 엑셀 업로드
            </button>
            <input ref={fileRef} type="file" accept=".xlsx,.xls,.csv"
              style={{display:"none"}} onChange={handleExcel}/>
          </div>
        </div>
        {formErr && <p style={{color:"#A32D2D",fontSize:12,marginTop:8}}>품목, 작업자, 수량을 모두 입력해 주세요.</p>}
        {uploadMsg && <p style={{fontSize:12,marginTop:8,color:uploadMsg.startsWith("✅")?"#3B6D11":"#A32D2D"}}>{uploadMsg}</p>}
      </div>

      {/* 검색 */}
      <div style={S.card}>
        <p style={S.secTitle}>🔍 기간별 작업자 조회</p>
        <div style={S.srchGrid}>
          <div><label style={S.lbl}>시작일</label><input ref={refSrchFrom} style={S.inp} type="date" value={srchFrom} onChange={e=>setSrchFrom(e.target.value)} onKeyDown={e=>e.key==="Enter"&&refSrchTo.current?.focus()}/></div>
          <div><label style={S.lbl}>종료일</label><input ref={refSrchTo} style={S.inp} type="date" value={srchTo} onChange={e=>setSrchTo(e.target.value)} onKeyDown={e=>e.key==="Enter"&&refSrchWorker.current?.focus()}/></div>
          <div><label style={S.lbl}>작업자(비우면 전체)</label><input ref={refSrchWorker} style={S.inp} type="text" placeholder="이름" value={srchWorker} onChange={e=>setSrchWorker(e.target.value)} onKeyDown={e=>e.key==="Enter"&&doSearch()}/></div>
          <div style={{display:"flex",alignItems:"flex-end"}}>
            <button style={S.btnSrch} onClick={doSearch}>📊 검색</button>
          </div>
        </div>
      </div>

      {/* 검색 결과 */}
      {report && (
        <div style={{...S.card, border:"0.5px solid #3B6D11"}}>
          <div style={{display:"flex",justifyContent:"space-between",alignItems:"center",marginBottom:12,flexWrap:"wrap",gap:6}}>
            <span style={{fontSize:14,fontWeight:600}}>📊 {report.from} ~ {report.to} 작업 현황 (총 {report.total}건)</span>
            <button style={S.btnClose} onClick={()=>setReport(null)}>✕ 닫기</button>
          </div>
          <div style={{overflowX:"auto"}}>
            <table style={{width:"100%",borderCollapse:"collapse",fontSize:13,minWidth:520}}>
              <thead>
                <tr style={{background:"#f8f7f3"}}>
                  <th style={S.rth} rowSpan={2}>작업자</th>
                  <th style={{...S.rth,textAlign:"center",borderBottom:"0.5px solid #e0e0dc",background:"#EAF3DE"}} colSpan={3}>✅ 완료 (정산)</th>
                  <th style={{...S.rth,textAlign:"center",borderBottom:"0.5px solid #e0e0dc",borderLeft:"1px solid #e0e0dc",background:"#FFFBEA"}} colSpan={3}>⏳ 미완료·일부완료</th>
                </tr>
                <tr style={{background:"#f8f7f3"}}>
                  <th style={{...S.rth,textAlign:"right",background:"#EAF3DE"}}>건수</th>
                  <th style={{...S.rth,textAlign:"right",background:"#EAF3DE"}}>수량</th>
                  <th style={{...S.rth,textAlign:"right",background:"#EAF3DE"}}>금액</th>
                  <th style={{...S.rth,textAlign:"right",borderLeft:"1px solid #e0e0dc",background:"#FFFBEA"}}>건수</th>
                  <th style={{...S.rth,textAlign:"right",background:"#FFFBEA"}}>수량</th>
                  <th style={{...S.rth,textAlign:"right",background:"#FFFBEA"}}>금액</th>
                </tr>
              </thead>
              <tbody>
                {reportRows.map(([w,v])=>(
                  <tr key={w} style={{borderBottom:"0.5px solid #ebebeb"}}>
                    <td style={{padding:"8px 12px",fontWeight:500}}>{w}</td>
                    <td style={{padding:"8px 12px",textAlign:"right",color:"#3B6D11",background:"#F3FAE8"}}>{v.doneCnt}건</td>
                    <td style={{padding:"8px 12px",textAlign:"right",color:"#3B6D11",background:"#F3FAE8"}}>{fmt(v.doneQty)}</td>
                    <td style={{padding:"8px 12px",textAlign:"right",color:"#3B6D11",background:"#F3FAE8",fontWeight:600}}>{v.doneAmount?fmt(v.doneAmount)+"원":"—"}</td>
                    <td style={{padding:"8px 12px",textAlign:"right",color:"#92600A",borderLeft:"1px solid #ebebeb",background:"#FFFDF0"}}>{v.etcCnt}건</td>
                    <td style={{padding:"8px 12px",textAlign:"right",color:"#92600A",background:"#FFFDF0"}}>{fmt(v.etcQty)}</td>
                    <td style={{padding:"8px 12px",textAlign:"right",color:"#bbb",background:"#FFFDF0"}}>—</td>
                  </tr>
                ))}
                <tr style={{fontWeight:600}}>
                  <td style={{padding:"8px 12px",background:"#f3fae8"}}>합계</td>
                  <td style={{padding:"8px 12px",textAlign:"right",color:"#3B6D11",background:"#EAF3DE"}}>{reportRows.reduce((s,[,v])=>s+v.doneCnt,0)}건</td>
                  <td style={{padding:"8px 12px",textAlign:"right",color:"#3B6D11",background:"#EAF3DE"}}>{fmt(reportRows.reduce((s,[,v])=>s+v.doneQty,0))}</td>
                  <td style={{padding:"8px 12px",textAlign:"right",color:"#3B6D11",background:"#EAF3DE"}}>{fmt(reportRows.reduce((s,[,v])=>s+v.doneAmount,0))}원</td>
                  <td style={{padding:"8px 12px",textAlign:"right",color:"#92600A",borderLeft:"1px solid #e0e0dc",background:"#FFFBEA"}}>{reportRows.reduce((s,[,v])=>s+v.etcCnt,0)}건</td>
                  <td style={{padding:"8px 12px",textAlign:"right",color:"#92600A",background:"#FFFBEA"}}>{fmt(reportRows.reduce((s,[,v])=>s+v.etcQty,0))}</td>
                  <td style={{padding:"8px 12px",textAlign:"right",color:"#bbb",background:"#FFFBEA"}}>—</td>
                </tr>
              </tbody>
            </table>
          </div>
        </div>
      )}

      {/* 탭 */}
      <div style={{display:"flex",gap:6,flexWrap:"wrap",marginBottom:12}}>
        {[["all","전체"],["pending","미완료"],["partial","일부완료"],["done","완료"]].map(([k,l])=>(
          <button key={k} onClick={()=>{setTab(k);setPage(1);}}
            style={tab===k ? S.tabOn : S.tabOff}>{l}</button>
        ))}
      </div>

      {/* 테이블 */}
      <p style={{fontSize:11,color:"#aaa",textAlign:"right",marginBottom:4}}>← 좌우로 밀어서 확인하세요</p>
      <div style={S.tblWrap}>
        <table style={{borderCollapse:"collapse",fontSize:13,minWidth:760,width:"100%"}}>
          <thead>
            <tr style={{background:"#f8f7f3"}}>
              <th style={{...S.th,width:40}}>No.</th>
              <th style={{...S.th,minWidth:90}}>품목</th>
              <th style={{...S.th,minWidth:75}}>작업자</th>
              <th style={{...S.th,width:52,textAlign:"center"}}>수량</th>
              <th style={{...S.th,width:76,textAlign:"right"}}>단가</th>
              <th style={{...S.th,width:84,textAlign:"right"}}>금액</th>
              <th style={{...S.th,width:96}}>지급일/시간</th>
              <th style={{...S.th,width:88}}>완료일</th>
              <th style={{...S.th,width:68,textAlign:"center"}}>상태</th>
              <th style={{...S.th,width:108,textAlign:"center"}}>처리</th>
              <th style={{...S.th,width:68,textAlign:"center"}}>작성자</th>
              <th style={{...S.th,width:42,textAlign:"center"}}>삭제</th>
            </tr>
          </thead>
          <tbody>
            {filtered.length===0 ? (
              <tr><td colSpan={11} style={{padding:36,textAlign:"center",color:"#aaa",fontSize:13}}>
                등록된 작업이 없습니다.
              </td></tr>
            ) : paginated.map((j,idx)=>(
              <tr key={j.fbKey} style={{background: j.status==="done"?"#F3FAE8":j.status==="partial"?"#FFFDF0":"#FEF6F6", borderBottom:"0.5px solid #ebebeb"}}>
                <td style={{padding:"9px 12px",color:"#aaa"}}>{idx+1}</td>
                <td style={{padding:"9px 12px",fontWeight:500,color:j.status==="done"?"#A32D2D":j.status==="partial"?"#92600A":"#1a1a1a",textDecoration:j.status==="done"?"line-through":"none",whiteSpace:"normal",minWidth:90}}>
                  {canEdit(j) ? <span style={S.editable} onClick={()=>openModal(j,"item")}>{j.item}</span> : j.item}
                </td>
                <td style={{padding:"9px 12px",color:j.status==="done"?"#A32D2D":j.status==="partial"?"#92600A":"#1a1a1a",textDecoration:j.status==="done"?"line-through":"none",whiteSpace:"normal",minWidth:75}}>
                  {canEdit(j) ? <span style={S.editable} onClick={()=>openModal(j,"worker")}>{j.worker}</span> : j.worker}
                </td>
                <td style={{padding:"9px 12px",textAlign:"center",color:j.status==="done"?"#A32D2D":j.status==="partial"?"#92600A":"#1a1a1a",textDecoration:j.status==="done"?"line-through":"none"}}>
                  {canEdit(j) ? <span style={S.editable} onClick={()=>openModal(j,"qty")}>{j.qty}</span> : j.qty}
                </td>
                <td style={{padding:"9px 12px",textAlign:"right",fontSize:12,color:j.status==="done"?"#A32D2D":j.status==="partial"?"#92600A":"#1a1a1a",textDecoration:j.status==="done"?"line-through":"none"}}>
                  {canEdit(j) ? <span style={S.editable} onClick={()=>openModal(j,"price")}>{j.price?fmt(j.price)+"원":"—"}</span> : (j.price?fmt(j.price)+"원":"—")}
                </td>
                <td style={{padding:"9px 12px",textAlign:"right",fontSize:12}}>
                  {j.price ? <strong>{fmt(Number(j.qty)*Number(j.price))}원</strong> : <span style={{color:"#bbb"}}>—</span>}
                </td>
                <td style={{padding:"9px 12px",fontSize:12,color:j.status==="done"?"#A32D2D":j.status==="partial"?"#92600A":"#1a1a1a",textDecoration:j.status==="done"?"line-through":"none"}}>
                  {(() => {
                    const parts = (j.date||"").split(" ");
                    return parts.length >= 2
                      ? <><span>{parts[0]}</span><br/><span style={{color:"#888",fontSize:11}}>{parts[1]}</span></>
                      : j.date;
                  })()}
                </td>
                <td style={{padding:"9px 12px",fontSize:12}}>{j.doneDate?<span style={{color:"#3B6D11"}}>{j.doneDate}</span>:<span style={{color:"#bbb"}}>—</span>}</td>
                <td style={{padding:"9px 12px",textAlign:"center"}}><Pill j={j}/></td>
                <td style={{padding:"9px 12px",textAlign:"center"}}><ActBtns j={j}/></td>
                <td style={{padding:"9px 12px",textAlign:"center"}}>
                  <span style={{fontSize:11,padding:"3px 8px",borderRadius:20,background:"#f0efeb",color:"#555",whiteSpace:"nowrap"}}>
                    {j.createdBy||"—"}
                  </span>
                </td>
                <td style={{padding:"9px 12px",textAlign:"center"}}>
                  {role==="master"
                    ? <button onClick={()=>deleteJob(j)} style={{background:"none",border:"none",cursor:"pointer",color:"#bbb",fontSize:18,padding:4}}>🗑</button>
                    : <span style={{color:"#ccc",fontSize:14}}>🔒</span>}
                </td>
              </tr>
            ))}
          </tbody>
        </table>
      </div>

      {/* 페이지네이션 */}
      {filtered.length > PAGE_SIZE && (
        <div style={{display:"flex",alignItems:"center",justifyContent:"center",gap:4,padding:"14px 0",flexWrap:"wrap"}}>
          <button onClick={()=>setPage(1)} disabled={page===1}
            style={{...S.pgBtn,opacity:page===1?0.3:1}}>«</button>
          <button onClick={()=>setPage(p=>Math.max(1,p-1))} disabled={page===1}
            style={{...S.pgBtn,opacity:page===1?0.3:1}}>‹</button>

          {Array.from({length:totalPages},(_,i)=>i+1)
            .filter(n=>n===1||n===totalPages||Math.abs(n-page)<=4)
            .reduce((acc,n,i,arr)=>{
              if(i>0&&n-arr[i-1]>1) acc.push("...");
              acc.push(n); return acc;
            },[])
            .map((n,i)=> n==="..."
              ? <span key={"e"+i} style={{padding:"0 4px",color:"#aaa",fontSize:13}}>…</span>
              : <button key={n} onClick={()=>setPage(n)}
                  style={{...S.pgBtn,
                    background:page===n?"#1a1a1a":"transparent",
                    color:page===n?"#fff":"#1a1a1a",
                    border:page===n?"none":"0.5px solid #ccc",
                    fontWeight:page===n?600:400,
                    minWidth:32}}>
                  {n}
                </button>
            )
          }

          <button onClick={()=>setPage(p=>Math.min(totalPages,p+1))} disabled={page===totalPages}
            style={{...S.pgBtn,opacity:page===totalPages?0.3:1}}>›</button>
          <button onClick={()=>setPage(totalPages)} disabled={page===totalPages}
            style={{...S.pgBtn,opacity:page===totalPages?0.3:1}}>»</button>

          <span style={{fontSize:12,color:"#888",marginLeft:8,display:"flex",alignItems:"center",gap:4}}>
            <input type="number" min={1} max={totalPages}
              defaultValue={page} key={page}
              onKeyDown={e=>{if(e.key==="Enter"){const v=parseInt(e.target.value);if(v>=1&&v<=totalPages)setPage(v);}}}
              style={{width:48,height:30,border:"0.5px solid #ccc",borderRadius:6,padding:"0 6px",fontSize:12,textAlign:"center",fontFamily:"inherit",outline:"none",boxSizing:"border-box"}}
            />
            <span>/ {totalPages}</span>
          </span>
        </div>
      )}

      {/* 범례 */}
      <div style={{display:"flex",gap:14,flexWrap:"wrap",padding:"2px 0 8px",fontSize:11}}>
        <span style={{color:"#A32D2D",display:"flex",alignItems:"center",gap:4}}><span style={{width:9,height:9,borderRadius:"50%",background:"#F7C1C1",display:"inline-block"}}/> 미완료</span>
        <span style={{color:"#6B4400",display:"flex",alignItems:"center",gap:4}}><span style={{width:0,height:0,borderLeft:"5px solid transparent",borderRight:"5px solid transparent",borderBottom:"9px solid #FAE588",display:"inline-block"}}/> 일부완료</span>
        <span style={{color:"#3B6D11",display:"flex",alignItems:"center",gap:4}}><span style={{width:9,height:9,borderRadius:"50%",background:"#C0DD97",display:"inline-block"}}/> 완료</span>
      </div>

      {/* 접속/활동 이력 모달 */}
      {showLog && (
        <div style={S.modalBg} onClick={e=>e.target===e.currentTarget&&setShowLog(false)}>
          <div style={{...S.modalBox, width:"min(560px,96vw)", maxHeight:"85vh", display:"flex", flexDirection:"column"}}>
            <div style={{display:"flex",justifyContent:"space-between",alignItems:"center",marginBottom:14}}>
              <h3 style={{fontSize:15,fontWeight:600}}>👁 관리자 이력</h3>
              <button style={S.btnCancel} onClick={()=>setShowLog(false)}>✕ 닫기</button>
            </div>
            {/* 탭 */}
            <div style={{display:"flex",gap:6,marginBottom:12}}>
              {[["activity","📋 작업이력"],["access","🔑 접속이력"]].map(([k,l])=>(
                <button key={k} onClick={()=>setLogTab(k)}
                  style={{fontSize:12,padding:"6px 14px",borderRadius:8,cursor:"pointer",fontFamily:"inherit",
                    background:logTab===k?"#1a1a1a":"transparent",
                    color:logTab===k?"#fff":"#1a1a1a",
                    border:logTab===k?"none":"0.5px solid #ccc"}}>
                  {l}
                </button>
              ))}
            </div>
            <div style={{overflowY:"auto",flex:1}}>
              {logTab === "activity" ? (
                <table style={{width:"100%",borderCollapse:"collapse",fontSize:13}}>
                  <thead>
                    <tr style={{background:"#f8f7f3",position:"sticky",top:0}}>
                      <th style={{padding:"8px 10px",textAlign:"left",fontWeight:500,color:"#888",fontSize:11,whiteSpace:"nowrap"}}>시각</th>
                      <th style={{padding:"8px 10px",textAlign:"left",fontWeight:500,color:"#888",fontSize:11,whiteSpace:"nowrap"}}>이름</th>
                      <th style={{padding:"8px 10px",textAlign:"left",fontWeight:500,color:"#888",fontSize:11,whiteSpace:"nowrap"}}>작업</th>
                      <th style={{padding:"8px 10px",textAlign:"left",fontWeight:500,color:"#888",fontSize:11}}>내용</th>
                    </tr>
                  </thead>
                  <tbody>
                    {activityLogs.length === 0
                      ? <tr><td colSpan={4} style={{padding:24,textAlign:"center",color:"#aaa"}}>작업 이력이 없습니다.</td></tr>
                      : activityLogs.map((log,i) => {
                          const actionColor =
                            log.action==="삭제" ? "#A32D2D" :
                            log.action==="완료처리" ? "#3B6D11" :
                            log.action==="일부완료처리" ? "#92600A" :
                            log.action==="엑셀업로드" ? "#1a56db" :
                            log.action==="등록" ? "#1a1a1a" : "#555";
                          return (
                            <tr key={log.key||i} style={{borderBottom:"0.5px solid #ebebeb",background:i===0?"#F0F7FF":"transparent"}}>
                              <td style={{padding:"7px 10px",fontSize:11,color:"#888",whiteSpace:"nowrap"}}>{log.time}</td>
                              <td style={{padding:"7px 10px",fontWeight:500,whiteSpace:"nowrap"}}>{log.name}</td>
                              <td style={{padding:"7px 10px",whiteSpace:"nowrap"}}>
                                <span style={{fontSize:11,padding:"2px 7px",borderRadius:20,fontWeight:500,
                                  background:actionColor==="#A32D2D"?"#FCEBEB":actionColor==="#3B6D11"?"#EAF3DE":actionColor==="#92600A"?"#FFFBEA":actionColor==="#1a56db"?"#E8F0FE":"#f0efeb",
                                  color:actionColor}}>
                                  {log.action}
                                </span>
                              </td>
                              <td style={{padding:"7px 10px",fontSize:12,color:"#444"}}>{log.detail}</td>
                            </tr>
                          );
                        })
                    }
                  </tbody>
                </table>
              ) : (
                <table style={{width:"100%",borderCollapse:"collapse",fontSize:13}}>
                  <thead>
                    <tr style={{background:"#f8f7f3",position:"sticky",top:0}}>
                      <th style={{padding:"8px 12px",textAlign:"left",fontWeight:500,color:"#888",fontSize:11}}>이름</th>
                      <th style={{padding:"8px 12px",textAlign:"left",fontWeight:500,color:"#888",fontSize:11}}>권한</th>
                      <th style={{padding:"8px 12px",textAlign:"left",fontWeight:500,color:"#888",fontSize:11}}>접속시각</th>
                    </tr>
                  </thead>
                  <tbody>
                    {accessLogs.length === 0
                      ? <tr><td colSpan={3} style={{padding:24,textAlign:"center",color:"#aaa"}}>접속 이력이 없습니다.</td></tr>
                      : accessLogs.map((log,i) => (
                        <tr key={log.key||i} style={{borderBottom:"0.5px solid #ebebeb",background:i===0?"#F0F7FF":"transparent"}}>
                          <td style={{padding:"8px 12px",fontWeight:500}}>{log.name}</td>
                          <td style={{padding:"8px 12px"}}>
                            <span style={{fontSize:11,padding:"2px 8px",borderRadius:20,fontWeight:500,
                              background:log.role==="master"?"#1a1a1a":"#f0efeb",
                              color:log.role==="master"?"#fff":"#555"}}>
                              {log.role==="master"?"👑 마스터":"👤 일반"}
                            </span>
                          </td>
                          <td style={{padding:"8px 12px",fontSize:12,color:"#666"}}>{log.time}</td>
                        </tr>
                      ))
                    }
                  </tbody>
                </table>
              )}
            </div>
          </div>
        </div>
      )}

      {/* 편집 모달 */}
      {modal && (
        <div style={S.modalBg} onClick={e=>e.target===e.currentTarget&&setModal(null)}>
          <div style={S.modalBox}>
            <h3 style={{fontSize:15,fontWeight:600,marginBottom:4}}>
              {modal.field==="qty"?"수량 수정":modal.field==="price"?"단가 수정":modal.field==="item"?"품목 수정":"작업자 수정"}
            </h3>
            <p style={{fontSize:12,color:"#888",marginBottom:16}}>{modal.sub}</p>
            <label style={S.lbl}>
              {modal.field==="qty"?"수량":modal.field==="price"?"단가 (원)":modal.field==="item"?"품목명":"작업자명"}
            </label>
            <input ref={modalRef}
              style={{...S.modalInp, textAlign:(modal.field==="item"||modal.field==="worker")?"left":"right"}}
              type={modal.field==="item"||modal.field==="worker"?"text":"number"}
              min={modal.field==="qty"?1:0} value={modalVal}
              onChange={e=>setModalVal(e.target.value)}
              onKeyDown={e=>{if(e.key==="Enter")saveModal();if(e.key==="Escape")setModal(null);}}/>
            <div style={{display:"flex",gap:8}}>
              <button style={S.btnSave} onClick={saveModal}>저장</button>
              <button style={S.btnCancel} onClick={()=>setModal(null)}>취소</button>
            </div>
          </div>
        </div>
      )}
    </div>
  );
}

// ── 스타일 상수 ────────────────────────────────────────────
const S = {
  wrap:        { maxWidth:1080, margin:"0 auto", padding:"16px 12px", fontFamily:"-apple-system,'Malgun Gothic',sans-serif", color:"#1a1a1a", minHeight:"100vh", background:"#f5f4f0" },
  loginOverlay:{ position:"fixed", inset:0, background:"#f5f4f0", zIndex:100, display:"flex", alignItems:"center", justifyContent:"center", padding:"0 16px" },
  loginBox:    { background:"#fff", border:"0.5px solid #e0e0dc", borderRadius:16, padding:"32px 28px", width:"min(340px,92vw)", margin:"0 auto" },
  loginLogo:   { display:"block", margin:"0 auto 20px", height:100, objectFit:"contain" },
  loginTitle:  { fontSize:16, fontWeight:600, color:"#888", textAlign:"center", marginBottom:18 },
  loginBtn:    { width:"100%", height:42, border:"none", borderRadius:8, background:"#1a1a1a", color:"#fff", fontSize:15, cursor:"pointer", fontWeight:500 },
  loginErr:    { fontSize:12, color:"#A32D2D", marginTop:8, textAlign:"center" },
  header:      { display:"flex", alignItems:"flex-start", justifyContent:"space-between", marginBottom:14, flexWrap:"wrap", gap:8 },
  headerLeft:  { display:"flex", alignItems:"center", gap:10 },
  headerLogo:  { height:32, objectFit:"contain" },
  h1:          { fontSize:18, fontWeight:700, margin:0 },
  sub:         { fontSize:11, color:"#888", marginTop:2 },
  headerRight: { display:"flex", alignItems:"center", gap:8, flexWrap:"wrap" },
  chipMaster:  { fontSize:12, padding:"4px 10px", borderRadius:8, fontWeight:500, background:"#1a1a1a", color:"#fff" },
  chipStaff:   { fontSize:12, padding:"4px 10px", borderRadius:8, fontWeight:500, background:"#f0efeb", color:"#555" },
  logoutBtn:   { fontSize:12, padding:"5px 10px", borderRadius:8, border:"0.5px solid #ccc", background:"transparent", color:"#888", cursor:"pointer" },
  badges:      { display:"flex", gap:6, flexWrap:"wrap", marginBottom:8 },
  badgeRed:    { fontSize:11, padding:"4px 10px", borderRadius:8, fontWeight:500, background:"#FCEBEB", color:"#A32D2D" },
  badgeYellow: { fontSize:11, padding:"4px 10px", borderRadius:8, fontWeight:500, background:"#FFFBEA", color:"#92600A" },
  badgeGreen:  { fontSize:11, padding:"4px 10px", borderRadius:8, fontWeight:500, background:"#EAF3DE", color:"#3B6D11" },
  badgeBlue:   { fontSize:11, padding:"4px 10px", borderRadius:8, fontWeight:500, background:"#E8F0FE", color:"#1a56db" },

  syncBar:     { fontSize:11, color:"#1a56db", textAlign:"right", marginBottom:6, minHeight:16 },
  card:        { background:"#fff", border:"0.5px solid #e0e0dc", borderRadius:12, padding:16, marginBottom:12 },
  secTitle:    { fontSize:13, fontWeight:500, color:"#888", marginBottom:12, display:"flex", alignItems:"center", gap:6 },
  todayChip:   { fontSize:11, background:"#f0efeb", color:"#666", padding:"2px 8px", borderRadius:6, marginLeft:4 },
  lbl:         { fontSize:12, color:"#888", display:"block", marginBottom:4 },
  inp:         { width:"100%", height:40, border:"0.5px solid #ccc", borderRadius:8, padding:"0 10px", fontSize:14, fontFamily:"inherit", outline:"none", background:"#fafaf8", WebkitAppearance:"none", appearance:"none", boxSizing:"border-box" },
  formGrid:    { display:"grid", gap:8, gridTemplateColumns:"repeat(auto-fit,minmax(120px,1fr))" },
  srchGrid:    { display:"grid", gap:8, gridTemplateColumns:"repeat(auto-fit,minmax(120px,1fr))" },
  btnAdd:      { height:40, padding:"0 18px", border:"none", borderRadius:8, background:"#1a1a1a", color:"#fff", fontSize:14, cursor:"pointer", whiteSpace:"nowrap", width:"100%" },
  btnSrch:     { height:40, padding:"0 18px", border:"none", borderRadius:8, background:"#2D6A4F", color:"#fff", fontSize:14, cursor:"pointer", whiteSpace:"nowrap", width:"100%" },
  btnExcel:    { height:40, padding:"0 18px", border:"0.5px solid #2D6A4F", borderRadius:8, background:"transparent", color:"#2D6A4F", fontSize:14, cursor:"pointer", whiteSpace:"nowrap" },
  btnClose:    { fontSize:12, padding:"5px 12px", borderRadius:8, border:"0.5px solid #ccc", background:"transparent", color:"#888", cursor:"pointer" },
  tblWrap:     { background:"#fff", border:"0.5px solid #e0e0dc", borderRadius:12, overflowX:"auto", WebkitOverflowScrolling:"touch", marginBottom:12 },
  th:          { padding:"9px 12px", textAlign:"left", fontWeight:500, color:"#888", fontSize:11, whiteSpace:"nowrap" },
  rth:         { padding:"8px 12px", textAlign:"left", fontWeight:500, color:"#888", fontSize:12 },
  tabOn:       { fontSize:12, padding:"7px 14px", borderRadius:8, cursor:"pointer", fontFamily:"inherit", background:"#1a1a1a", color:"#fff", border:"none" },
  tabOff:      { fontSize:12, padding:"7px 14px", borderRadius:8, cursor:"pointer", fontFamily:"inherit", background:"transparent", color:"#1a1a1a", border:"0.5px solid #ccc" },
  editable:    { cursor:"pointer", borderBottom:"1px dashed #bbb", paddingBottom:1 },
  modalBg:     { position:"fixed", inset:0, background:"rgba(0,0,0,0.25)", zIndex:200, display:"flex", alignItems:"center", justifyContent:"center" },
  modalBox:    { background:"#fff", borderRadius:14, padding:"24px 24px 20px", width:"min(300px,88vw)", boxShadow:"0 8px 32px rgba(0,0,0,0.14)" },
  modalInp:    { width:"100%", height:44, border:"0.5px solid #ccc", borderRadius:8, padding:"0 12px", fontSize:20, fontFamily:"inherit", outline:"none", background:"#fafaf8", textAlign:"right", marginBottom:14, boxSizing:"border-box" },
  btnSave:     { flex:1, height:40, border:"none", borderRadius:8, background:"#1a1a1a", color:"#fff", fontSize:14, cursor:"pointer", fontWeight:500 },
  btnCancel:   { flex:1, height:40, border:"0.5px solid #ccc", borderRadius:8, background:"transparent", color:"#888", fontSize:14, cursor:"pointer" },
  pgBtn:       { height:30, minWidth:30, padding:"0 6px", border:"0.5px solid #ccc", borderRadius:6, background:"transparent", color:"#1a1a1a", fontSize:13, cursor:"pointer", fontFamily:"inherit" },
};
