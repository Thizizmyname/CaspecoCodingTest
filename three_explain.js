 //****************************************************************************************************
 //
 //     3      Vad gör funktionen?
 //
 //**************************************************************************************************** 


 function explainThisFunc() {
	var self = this
	var groupsToRender = []
	var hasEditPermissions = this.hasPermission('booking.update')
	var actions = []

	if (hasEditPermissions) {
		actions.push(
			<Button
				id="addBooking"
				preset="addPerson"
				onClick={this.handleNewBookingClick}
				disabled={!hasEditPermissions}
			/>,
		)
	}

	var allLoaded =
		isLoaded(this.state.bookingsResult) &&
		isLoaded(this.state.bookedArticlesResult) &&
		isLoaded(this.state.currentUserResult) &&
		isLoaded(this.state.timePeriodStatisticsResult) &&
		isLoaded(this.state.tablesResult) &&
		isLoaded(this.state.tableArticlesResult)

	if (this.state.selectedUnit && allLoaded) {
		var tables = this.state.tablesResult.value

		var groups = this.state.selectedUnit.articleGroups

		var filteredGroups = groups.filter(
			(articleGroup) =>
				tables.findIndex((t) => t.articleGroupId === articleGroup.id) > -1,
		)

		var unspecifiedGroup = ArticleGroup.Record({
			articleGroupId: '0',
			name: this.t('booking.hasNoTables'),
		})

		filteredGroups = filteredGroups.push(unspecifiedGroup)

		filteredGroups.forEach(function (group) {
			var bookings = Immutable.List()

			self.state.bookingsResult.value.forEach(function (booking) {
				booking = booking.set(
					'articles',
					BookingStore.getBookedArticlesForBookingResult(
						booking.id,
						self.state.date,
						self.state.selectedUnit.id,
					).value,
				)

				var articlesInGroup = booking.articles.filter(
					(a) => a.articleGroupId === group.id,
				)

				if (
					articlesInGroup.count() > 0 ||
					(booking.articles.count() < 1 && group === unspecifiedGroup)
				) {
					bookings = bookings.push(booking)
				}
			})

            bookings = bookings.sort(function (a, b) {
				var moment1 = Moment(a.start, 'YYYY-MM-DDTHH:mm')

				var moment2 = Moment(b.start, 'YYYY-MM-DDTHH:mm')

				return moment1.diff(moment2, 'minutes')
			})


			var bookingsToRender = []
			bookings.forEach(function (booking) {
				bookingsToRender.push(
					<BookingListItem
						booking={booking}
						editBooking={self.state.editBooking}
						contacts={self.state.contacts}
						onChange={self.onChange}
						onClick={self.onClick}
					/>,
				)
			})

			groupsToRender.push(
				<BookingGroup
					key={group.name}
					title={group.name}
					object={group}
					editBooking={self.state.editBooking}
					contacts={self.state.contacts}
					expanded={false}
				>
					{bookingsToRender}
				</BookingGroup>,
			)
		})
	}


	var hasValidDaynote = this.hasValidDaynote()
	var daynote = this.getDaynoteText()

	var dateControl = (
		<div className="bookingView_header">
			<ToolBar wrapping="nowrap" className="bookingView_header_toolBar">
				<Datepicker
					id="booking.booking.date"
					value={this.state.date}
					fieldName={'date'}
					format="ddd Do MMM"
					canToggle={true}
					onChange={this.onChange}
				/>

				{allLoaded ?
					<ui.Select
						key="booking.timePeriods"
						id="booking.timePeriods"
						items={this.state.timePeriodStatisticsResult.value}
						dataValueField="name"
						dataTextField="text"
						value={this.state.selectedTimePeriodStatistics}
						fieldName="timePeriod"
						onChange={this.handleTimePeriodChange}
					/>
				    : null}

                {allLoaded && hasValidDaynote ? 
                    <ui.Icon
						preset="daynote"
						onClick={this.handleToggleDaynote}
						className="bookingView_header_daynote_icon"
					/>
				    : null}
			</ToolBar>

            {allLoaded && hasValidDaynote && this.state.daynoteExpanded ?
                <div className="bookingView_header_daynote_text"> {daynote} </div>
			    : null}
		</div>
	);

	return (
		<ui.Wrap className="bookingView">
			<ColumnLayout
				activeColumn={this.state.activeColumn}
				onActiveChanged={this.handleActiveColumnChange}
			>
				<Column
					sidebar={true}
					canGoLeft={true}
					name="PickUnit"
					title={this.t('booking.unit.pickUnit')}
				>
					<VerticalLayout>
						{allLoaded ?
							<ui.TableView
								id={'units'}
								items={this.state.unitsResult.value}
								itemKey={'id'}
								onSelectItem={this.handleSelectUnit}
								selectedItem={null}
								searchBar={false}
								columns={[
									{
										name: 'units',
										type: 'header',
										title: this.t('booking.booking.time'),
										value: 'name',
									},
								]}
								clickWillTransition={true}
							/>
						    : <ui.FlashBox type="loading" />
						}
					</VerticalLayout>
				</Column>

				<Column
					sidebar={true}
					canGoLeft={true}
					name="PickBooking"
					title={this.getUnitTitle(this.state.selectedUnit)}
					actionComponents={actions}
					controls={dateControl}
				>
					
					<VerticalLayout>
                        {allLoaded ? groupsToRender
                            : <ui.FlashBox type="loading" />}
					</VerticalLayout>
				</Column>
			</ColumnLayout>

			{this.state.newBooking && hasEditPermissions ?
				<EditBooking
					createMode       = {true}
					unit             = {this.state.selectedUnit}
					initialStartDate = {Moment(this.state.startFilter)}
					closeModal       = {this.handleCloseModal}
				/>
			: false}

			{this.state.editBooking ?
				<EditBooking
					createMode={false}
					unit={this.state.selectedUnit}
					booking={this.state.editBooking}
					closeModal={this.handleCloseModal}
				/>
			: false}
		</ui.Wrap>
	)
}