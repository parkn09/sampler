<?php
App::import('Vendor', 'ExcelWriter', array('file' => 'excelwriter.inc.php'));

class SheetsController extends AppController {

	var $name = 'Sheets';

	function index() {
		$this->Sheet->recursive = 0;
		$conditions = array();
		$conditions['Sheet.store_id'] = $this->Cookie->read('store_id');
		$conditions['Sheet.active'] = 1;
		$this->paginate = array('conditions'=>$conditions, 'fields'=>array('name', 'id', 'store_id', 'dateOf', 'tank_id'),'order'=>'Sheet.dateOf DESC', 'limit'=>15);
		$this->set('tanks', $this->Sheet->Tank->find('list', array('conditions'=>array('Tank.store_id'=>$this->Cookie->read("store_id")))));
		$this->set('sheets', $this->paginate());
	}


	function view($id = null) {
		if (!$id) {
			$this->Session->setFlash(__('Invalid sheet', true));
			$this->redirect(array('action' => 'index'));
		}
		$store_id = $this->Cookie->read('store_id');
		$sheet=$this->Sheet->find('first', array('conditions'=>array('Sheet.id'=>$id, 'Sheet.store_id'=>$store_id, 'Sheet.active'=>1)));
		if(!empty($sheet)){
			$begin = $sheet['Sheet']['dateOf'];
			$last = date('Y-m-t', strtotime($begin));
			$conditions = array();
			$conditions['Day.active']=1;
			$conditions['Day.store_id'] = $store_id;
			$conditions['Day.tank_id'] = $sheet['Sheet']['tank_id'];
			$conditions['Day.dateOf >='] = $begin;
			$conditions['Day.dateOf <='] = $last;
			$days = $this->Sheet->Store->Day->find('all', array('conditions'=>$conditions, 'recursive'=>-1, 'order'=>'Day.dateOf ASC'));
			$this->set('days',$days);

		}
		$this->set('sheet', $sheet);
	}

	//a function to convert data from a table to an Excel Sheet using a plugin called ExcelWriter
	function convertto($id){
	if(isset($this->params['url']['isajax'])){
		$this->autoRender = false;
		$this->layout = 'ajax';
	}
		$store_id = $this->Cookie->read('store_id');
		$sheet=$this->Sheet->find('first', array('conditions'=>array('Sheet.id'=>$id, 'Sheet.store_id'=>$store_id, 'Sheet.active'=>1)));
		if(empty($sheet)){
			if(isset($this->params['url']['isajax'])) return false;
			$this->Session->setFlash('Sorry, but that sheet doesn\'t exist yet. Please create a sheet!');
			$this->redirect(array('action'=>'index'));
		}
			$begin = $sheet['Sheet']['dateOf'];
			$last = date('Y-m-t', strtotime($begin));
			$conditions = array();
			$conditions['Day.active']=1;
			$conditions['Day.store_id'] = $store_id;
			$conditions['Day.tank_id'] = $sheet['Sheet']['tank_id'];
			$conditions['Day.dateOf >='] = $begin;
			$conditions['Day.dateOf <='] = $last;
			$days = $this->Sheet->Store->Day->find('all', array('conditions'=>$conditions, 'recursive'=>-1, 'order'=>'Day.dateOf ASC'));
			$name = $sheet['Sheet']['name'].' for '.$sheet['Store']['name'].' on '.strtotime('now').'.xls';
			$file = "excelfiles/".$name;
			$excel = new ExcelWriter($file);
			$excel->writeRow();
			$excel->writeRow();
			$excel->writeCol($sheet['Sheet']['name'].' for '.$sheet['Store']['name']);
			$excel->writeRow(); $excel->writeRow();
			$excel->writeCol('Date'); $excel->writeCol('Beginning Stick');
			$excel->writeCol('Delivered'); $excel->writeCol('Pumped'); $excel->writeCol('Ending Stick');
			$excel->writeCol('Ending Book'); $excel->writeCol('Over/Under'); $excel->writeCol('Water (inches)');
			$stats = array();
			$stats['Pumped'] = 0;
			$stats['Delivered'] = 0;
			$stats['Over']=0;
			$stats['Water'] = 0;
			for($x=0;$x<10;$x++){
				//$excel->writeRow();
				if(isset($days[$x]['Day']['id'])){
					if($x==0) $first = date('Y-m-d', strtotime($days[$x]['Day']['dateOf']));
					$stuff = array();
					$stuff[] = date('m/d/Y', strtotime($days[$x]['Day']['dateOf']));
					$stuff[] = date($days[$x]['Day']['stick']);
					$stats['Delivered'] +=$days[$x]['Day']['delivered'];
					$pumped = 0;
					if(isset($days[$x+1]['Day'])){
						$pumped = $days[$x+1]['Day']['start'];
						$stats['Pumped']+=$pumped;
					}
					$stuff[] = $pumped;	
					$book = $days[$x]['Day']['stick']+$days[$x]['Day']['delivered']-$pumped;
					$stuff[] = $book;
					$over = 0;
					if(isset($days[$x+1]['Day'])){
						$stuff[] = $days[$x+1]['Day']['stick'];
						$over = $days[$x+1]['Day']['stick']-$book;
					}
					$stuff[] = $over;
					$stats['Over']+=$over;
					$stuff[] = $days[$x]['Day']['water'];
					$stats['Water']+=$days[$x]['Day']['water'];
					$excel->writeLine($stuff);				
				}else{
					$excel->writeRow();
					$excel->writeCol(date('m-d-Y', strtotime($first."+ $x days")));
				}
			}
			$excel->close();
				$this->data = array();
				$this->data['Thefile']['name'] =$name;
				$this->data['Thefile']['sheet_id'] = $id;
				$this->data['Thefile']['store_id'] = $this->Cookie->read('store_id');
				$this->data['Thefile']['type'] =1;
				$this->data['Thefile']['active'] =1;
				if($this->Sheet->Thefile->save($this->data['Thefile'])){
					$this->data['Sheet']['id'] = $id;
					$this->data['Sheet']['thefile_id'] = $this->Sheet->Thefile->id;
					$this->Sheet->updateAll(array('Sheet.thefile_id'=>$this->Sheet->Thefile->id), array('Sheet.id'=>$id));
					return $this->Sheet->Thefile->id;
				}			
			return false;
		
	}

	function add() {
		if(isset($this->params['url']['isajax'])){
			$this->autoRender = false;
			$this->layout = 'ajax';
			$tank = $this->Sheet->Tank->find('first', array('conditions'=>array('Tank.id'=>$_POST['tank_id'], 'Tank.store_id'=>$this->Cookie->read('store_id')), 'recursive'=>-1, 'fields'=>array("id", "store_id", "name")));
			$this->data['Sheet']['tank_id'] = $_POST['tank_id'];
			$this->data['Sheet']['store_id'] = $this->Cookie->read('store_id');
			$this->data['Sheet']['dateOf'] = $_POST['year'].'-'.$_POST['month'].'-01';
			$this->data['Sheet']['name'] = $tank['Tank']['name'].' '.date('F', strtotime($this->data['Sheet']['dateOf'])).' '.$_POST['year'];
			if($this->Sheet->save($this->data)){
				$data = array($this->data['Sheet']['name'], $this->Sheet->id);
				return json_encode($data);
			}
			return false;
		}
		if (!empty($this->data)) {
			$store_id = $this->Cookie->read("store_id");
			$this->Sheet->create();
			$this->data['Sheet']['store_id'] = $store_id;
			$this->data['Sheet']['dateOf'] = $this->data['Sheet']['year'].'-'.$this->data['Sheet']['month'].'-01';
			if ($this->Sheet->save($this->data)) {
				$this->Session->setFlash(__('The sheet has been saved', true));
				$this->redirect(array('action' => 'index'));
			} else {
				$this->Session->setFlash(__('The sheet could not be saved. Please, try again.', true));
			}
		}
		$tanks = $this->Sheet->Tank->find('list');
		$stores = $this->Sheet->Store->find('list');
		$users = $this->Sheet->User->find('list');
		$this->set(compact('tanks', 'stores', 'users'));
	}

	function edit($id = null) {
		if (!$id && empty($this->data)) {
			$this->Session->setFlash(__('Invalid sheet', true));
			$this->redirect(array('action' => 'index'));
		}
		if (!empty($this->data)) {
			if ($this->Sheet->save($this->data)) {
				$this->Session->setFlash(__('The sheet has been saved', true));
				$this->redirect(array('action' => 'index'));
			} else {
				$this->Session->setFlash(__('The sheet could not be saved. Please, try again.', true));
			}
		}
		if (empty($this->data)) {
			$this->data = $this->Sheet->read(null, $id);
		}
		$tanks = $this->Sheet->Tank->find('list');
		$stores = $this->Sheet->Store->find('list');
		$users = $this->Sheet->User->find('list');
		$this->set(compact('tanks', 'stores', 'users'));
	}

	function delete($id = null) {
		if (!$id) {
			$this->Session->setFlash(__('Invalid id for sheet', true));
			$this->redirect(array('action'=>'index'));
		}
		if ($this->Sheet->delete($id)) {
			$this->Session->setFlash(__('Sheet deleted', true));
			$this->redirect(array('action'=>'index'));
		}
		$this->Session->setFlash(__('Sheet was not deleted', true));
		$this->redirect(array('action' => 'index'));
	}
}
?>
